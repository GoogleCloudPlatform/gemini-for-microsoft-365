# Copyright 2026 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import os
import sys
import time
import json
import logging
from datetime import datetime, timezone
from typing import List, Optional, Dict, Any, Tuple
import urllib.request
import requests
from fastapi import FastAPI, HTTPException, Request, Depends, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel, Field
import httpx
import jwt
from dotenv import load_dotenv

# Load local environment variables from root or local .env file
root_env = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".env"))
if os.path.exists(root_env):
    load_dotenv(root_env)
load_dotenv()

# ---------------------------------------------------------------------------
# Native Google Cloud Structured JSON Logging Configuration
# ---------------------------------------------------------------------------
VERBOSE_LOGGING = os.environ.get("VERBOSE_LOGGING", "false").lower() in ("true", "1", "yes")
LOG_LEVEL_ENV = os.environ.get("LOG_LEVEL", "DEBUG" if VERBOSE_LOGGING else "INFO").upper()

# Standard GCP Severity Mapping
GCP_SEVERITY_MAP = {
    "DEBUG": "DEBUG",
    "INFO": "INFO",
    "WARNING": "WARNING",
    "ERROR": "ERROR",
    "CRITICAL": "CRITICAL",
}


class GCPStructuredJsonFormatter(logging.Formatter):
    """
    Formats log records as structured JSON payloads natively parsed by Google Cloud Logging.
    Merges custom 'extra' parameters directly into the JSON entry for rich indexing and filtering.
    """
    def format(self, record: logging.LogRecord) -> str:
        log_payload = {
            "severity": GCP_SEVERITY_MAP.get(record.levelname, "INFO"),
            "message": record.getMessage(),
            "time": datetime.now(timezone.utc).isoformat(),
            "logger": record.name,
            "logging.googleapis.com/labels": {
                "service": "auth-proxy",
                "version": "1.0.0",
                "verbose_mode": str(VERBOSE_LOGGING).lower(),
            }
        }

        # Include exception tracebacks if present
        if record.exc_info:
            log_payload["exception"] = self.formatException(record.exc_info)

        # Merge extra structured attributes passed to logger calls
        extra_fields = {
            k: v for k, v in record.__dict__.items()
            if k not in (
                "args", "asctime", "created", "exc_info", "exc_text", "filename",
                "funcName", "id", "levelname", "levelno", "lineno", "module",
                "msecs", "message", "msg", "name", "pathname", "process",
                "processName", "relativeCreated", "stack_info", "thread", "threadName"
            )
        }
        if extra_fields:
            log_payload["structured_context"] = extra_fields

        return json.dumps(log_payload, default=str)


# Initialize Root & App Loggers with GCP JSON Formatter
handler = logging.StreamHandler(sys.stdout)
handler.setFormatter(GCPStructuredJsonFormatter())

logger = logging.getLogger("auth-proxy")
logger.setLevel(getattr(logging, LOG_LEVEL_ENV, logging.INFO))
logger.handlers = [handler]
logger.propagate = False

# Also set uvicorn access logger if running standalone
logging.getLogger("uvicorn.access").handlers = [handler]

logger.info(
    "auth-proxy logger initialized",
    extra={
        "log_level": LOG_LEVEL_ENV,
        "verbose_logging": VERBOSE_LOGGING,
        "gcp_logging_mode": "structured_json"
    }
)

# ---------------------------------------------------------------------------
# Microsoft Entra ID (Azure AD) SSO Configuration
# ---------------------------------------------------------------------------
JWKS_URL = os.environ.get(
    "ENTRA_JWKS_URL", 
    "https://login.microsoftonline.com/common/discovery/v2.0/keys"
)
jwks_client = jwt.PyJWKClient(JWKS_URL)

# Cloud Run Deployment Project & Service Account
PROJECT_ID = os.environ.get("PROJECT_ID", "")
REGION = os.environ.get("REGION", os.environ.get("GCP_REGION", "us-central1"))
SERVICE_ACCOUNT = os.environ.get("SERVICE_ACCOUNT", "")

# Gemini Enterprise (Discovery Engine) Target
GE_GCP_PROJECT_ID = os.environ.get("GE_GCP_PROJECT_ID") or os.environ.get("GCP_PROJECT_ID") or PROJECT_ID
GE_GCP_LOCATION = os.environ.get("GE_GCP_LOCATION") or os.environ.get("GCP_LOCATION", "us")
STREAM_ASSIST_ENDPOINT_LOCATION = os.environ.get("STREAM_ASSIST_ENDPOINT_LOCATION") or GE_GCP_LOCATION
GEMINI_ENTERPRISE_APP_ID = os.environ.get("GEMINI_ENTERPRISE_APP_ID", "")
ENTERPRISE_COLLECTION_ID = os.environ.get("ENTERPRISE_COLLECTION_ID") or os.environ.get("GEMINI_ENTERPRISE_COLLECTION_ID", "default_collection")
ENTERPRISE_ASSISTANT_ID = os.environ.get("ENTERPRISE_ASSISTANT_ID") or os.environ.get("GEMINI_ENTERPRISE_ASSISTANT_ID", "default_assistant")

# Microsoft Entra ID & Google OAuth Client IDs
MICROSOFT_ENTRA_APP_ID = os.environ.get("MICROSOFT_ENTRA_APP_ID", "")
MICROSOFT_ENTRA_TENANT_ID = os.environ.get("MICROSOFT_ENTRA_TENANT_ID", "")
GOOGLE_OAUTH_CLIENT_ID = os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "")

ENTRA_APP_ID = MICROSOFT_ENTRA_APP_ID
ENTRA_TENANT_ID = MICROSOFT_ENTRA_TENANT_ID
GCP_PROJECT_ID = GE_GCP_PROJECT_ID
GCP_LOCATION = GE_GCP_LOCATION
REQUIRE_ENTRA_AUTH = os.environ.get("REQUIRE_ENTRA_AUTH", "true").lower() in ("true", "1", "yes")
DOWNSTREAM_BACKEND_URL = os.environ.get("DOWNSTREAM_BACKEND_URL", "").rstrip("/")
USER_AUTH_MODE = os.environ.get("USER_AUTH_MODE", "auto").lower()
WIF_AUDIENCE = os.environ.get("WIF_AUDIENCE", "")
WIF_PROVIDER_NAME = os.environ.get("WIF_PROVIDER_NAME", "entra-id-oidc-pool-provider")

logger.info(
    "Entra ID & Downstream configuration loaded",
    extra={
        "require_entra_auth": REQUIRE_ENTRA_AUTH,
        "entra_app_id_configured": bool(ENTRA_APP_ID),
        "entra_app_id_preview": f"{ENTRA_APP_ID[:8]}..." if ENTRA_APP_ID else "NONE",
        "downstream_backend_configured": bool(DOWNSTREAM_BACKEND_URL),
        "downstream_backend_url": DOWNSTREAM_BACKEND_URL or "NONE (Isolated Dev Mode)",
        "gcp_project_id": GCP_PROJECT_ID,
        "gcp_location": GCP_LOCATION,
        "user_auth_mode": USER_AUTH_MODE,
        "wif_audience_override": WIF_AUDIENCE or "AUTO_DETECT",
        "jwks_url": JWKS_URL
    }
)

# In-memory cache for Discovery Engine aclConfig { "project_id:location": { "data": ..., "expires_at": ... } }
_ACL_CONFIG_CACHE: Dict[str, Dict[str, Any]] = {}

def get_gcp_access_token() -> Optional[str]:
    """Obtains a Google Cloud OAuth access token using Application Default Credentials (ADC)."""
    try:
        import google.auth
        from google.auth.transport.requests import Request as GoogleAuthRequest
        creds, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
        auth_req = GoogleAuthRequest()
        creds.refresh(auth_req)
        return creds.token
    except Exception as e:
        logger.debug(f"Could not obtain GCP ADC access token: {e}")
        return None

def get_discovery_engine_acl_config(project_id: str, location: str = "global") -> Dict[str, Any]:
    """
    Queries Google Cloud Discovery Engine project-level aclConfig to determine if
    the project uses Google Identity ('GSUITE') or Workforce Identity Federation ('THIRD_PARTY').
    Caches results in memory for 10 minutes to minimize API latency.
    """
    cache_key = f"{project_id}:{location}"
    now = time.time()
    if cache_key in _ACL_CONFIG_CACHE:
        entry = _ACL_CONFIG_CACHE[cache_key]
        if now < entry.get("expires_at", 0):
            return entry.get("data", {})

    if project_id == "agentspace-wif":
        return {
            "idpConfig": {
                "idpType": "THIRD_PARTY",
                "externalIdpConfig": {
                    "workforcePoolName": "locations/global/workforcePools/ca-entra-id-oidc-pool"
                }
            }
        }

    token = get_gcp_access_token()
    if not token:
        logger.warning(
            "Cannot query Discovery Engine aclConfig: GCP ADC token not available.",
            extra={"project_id": project_id, "location": location}
        )
        return {"idpConfig": {"idpType": "GSUITE"}}

    base_domain = "discoveryengine.googleapis.com" if location == "global" else f"{location}-discoveryengine.googleapis.com"
    url = f"https://{base_domain}/v1/projects/{project_id}/locations/{location}/aclConfig"

    try:
        resp = requests.get(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "X-Goog-User-Project": project_id,
                "Content-Type": "application/json"
            },
            timeout=5
        )
        if resp.status_code == 200:
            data = resp.json()
            _ACL_CONFIG_CACHE[cache_key] = {
                "data": data,
                "expires_at": now + 600  # 10 minutes TTL
            }
            logger.info(
                f"Discovered Gemini Enterprise Identity Provider settings for project '{project_id}': idpType='{data.get('idpConfig', {}).get('idpType')}'",
                extra={
                    "project_id": project_id,
                    "location": location,
                    "idp_config": data.get("idpConfig", {})
                }
            )
            return data
        else:
            logger.warning(
                f"Discovery Engine aclConfig returned HTTP {resp.status_code} for project '{project_id}'. Defaulting to GSUITE.",
                extra={"project_id": project_id, "location": location, "status_code": resp.status_code}
            )
            return {"idpConfig": {"idpType": "GSUITE"}}
    except Exception as e:
        logger.warning(
            f"Failed to query Discovery Engine aclConfig for project '{project_id}': {e}. Defaulting to GSUITE.",
            extra={"project_id": project_id, "location": location, "error": str(e)}
        )
        return {"idpConfig": {"idpType": "GSUITE"}}


def exchange_entra_jwt_for_wif_token(
    entra_jwt: str, 
    workforce_pool_name: str, 
    provider_name: Optional[str] = None
) -> Optional[str]:
    """
    Exchanges a Microsoft Entra ID JWT Bearer token for a Google Cloud Workforce Identity
    Federation (WIF) access token via Google Security Token Service (STS).
    """
    if not entra_jwt:
        return None

    clean_pool = workforce_pool_name.lstrip("/")
    
    # Candidate providers to attempt
    if WIF_AUDIENCE:
        candidate_audiences = [WIF_AUDIENCE]
    elif provider_name:
        candidate_audiences = [f"//iam.googleapis.com/{clean_pool}/providers/{provider_name}"]
    else:
        # Try default pool provider
        candidate_providers = [
            WIF_PROVIDER_NAME,
            "entra-id-oidc-pool-provider"
        ]
        # Deduplicate while preserving order
        seen = set()
        deduped_providers = []
        for p in candidate_providers:
            if p and p not in seen:
                seen.add(p)
                deduped_providers.append(p)
        candidate_audiences = [f"//iam.googleapis.com/{clean_pool}/providers/{p}" for p in deduped_providers]

    sts_url = "https://sts.googleapis.com/v1/token"
    start_t = time.time()

    for audience in candidate_audiences:
        sts_payload = {
            "grant_type": "urn:ietf:params:oauth:grant-type:token-exchange",
            "audience": audience,
            "scope": "https://www.googleapis.com/auth/cloud-platform",
            "requested_token_type": "urn:ietf:params:oauth:token-type:access_token",
            "subject_token": entra_jwt,
            "subject_token_type": "urn:ietf:params:oauth:token-type:jwt"
        }

        try:
            res = requests.post(
                sts_url,
                data=sts_payload,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                timeout=8
            )
            latency_ms = round((time.time() - start_t) * 1000, 2)

            if res.status_code == 200:
                token_data = res.json()
                access_token = token_data.get("access_token")
                logger.info(
                    f"Successfully exchanged Entra ID token for Google WIF access token ({latency_ms}ms) via {audience}",
                    extra={
                        "sts_status": 200,
                        "sts_audience": audience,
                        "token_type": token_data.get("token_type"),
                        "expires_in": token_data.get("expires_in"),
                        "latency_ms": latency_ms
                    }
                )
                return access_token
            else:
                logger.warning(
                    f"Google STS token exchange failed for audience '{audience}' (HTTP {res.status_code}): {res.text}",
                    extra={
                        "sts_status": res.status_code,
                        "sts_audience": audience,
                        "sts_error": res.text[:500],
                        "latency_ms": latency_ms
                    }
                )
        except Exception as e:
            logger.error(
                f"Exception during Google STS token exchange for audience '{audience}': {e}",
                extra={"error": str(e), "sts_audience": audience},
                exc_info=True
            )

    return None


_user_token_cache: Dict[str, Tuple[str, float]] = {}
_oauth3_user_token_cache: Dict[str, Tuple[str, float]] = {}


def _user_cache_keys(user: Optional["AuthenticatedUser"]) -> List[str]:
    if not user:
        return []
    keys: List[str] = []
    for raw in (user.oid, user.email, user.user_id, user.sub):
        if raw and str(raw).strip():
            norm = str(raw).strip().lower()
            if norm not in keys:
                keys.append(norm)
    return keys


def cache_oauth3_google_token(
    user: Optional["AuthenticatedUser"],
    google_token: Optional[str],
    expires_in: int = 3600
) -> None:
    keys = _user_cache_keys(user)
    if not keys:
        return
    if not google_token or not str(google_token).strip():
        for k in keys:
            _oauth3_user_token_cache.pop(k, None)
        return
    clean_tok = str(google_token).strip()
    ttl = max(min(int(expires_in or 3600), 3600) - 120, 60)
    exp_ts = time.time() + ttl
    for k in keys:
        existing = _oauth3_user_token_cache.get(k)
        # Preserve existing expiry if same token is re-seen via request header
        if existing and existing[0] == clean_tok and existing[1] > time.time() + 60:
            continue
        _oauth3_user_token_cache[k] = (clean_tok, exp_ts)


def get_cached_oauth3_google_token(user: Optional["AuthenticatedUser"]) -> Tuple[Optional[str], int]:
    now = time.time()
    for k in _user_cache_keys(user):
        entry = _oauth3_user_token_cache.get(k)
        if entry:
            tok, exp_ts = entry
            remaining = int(exp_ts - now)
            if remaining > 60 and tok:
                return tok, remaining
            _oauth3_user_token_cache.pop(k, None)
    return None, 0

def mint_user_google_token_via_dwd(user_email: str) -> Optional[str]:
    """
    Automatically mints a short-lived Google OAuth2 Access Token for the given user email
    using Service Account Domain-Wide Delegation (DWD) and the IAM Credentials API.
    Cached for token lifetime.
    """
    if not user_email or "@" not in user_email:
        return None

    now = time.time()
    # Check cache (refresh if less than 5 minutes remaining)
    if user_email in _user_token_cache:
        cached_tok, exp = _user_token_cache[user_email]
        if exp - now > 300:
            return cached_tok

    try:
        from google.auth import default
        from google.auth.transport.requests import Request as GoogleAuthRequest
        import google.auth

        sa_email = SERVICE_ACCOUNT or (f"gemini-office365-sa@{PROJECT_ID}.iam.gserviceaccount.com" if PROJECT_ID else "")
        iat = int(now)
        exp_time = iat + 3600

        jwt_payload = {
            "iss": sa_email,
            "sub": user_email,
            "aud": "https://oauth2.googleapis.com/token",
            "scope": "https://www.googleapis.com/auth/cloud-platform https://www.googleapis.com/auth/drive.readonly",
            "iat": iat,
            "exp": exp_time
        }

        credentials, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
        auth_req = GoogleAuthRequest()
        credentials.refresh(auth_req)
        gcp_token = credentials.token

        sign_url = f"https://iamcredentials.googleapis.com/v1/projects/-/serviceAccounts/{sa_email}:signJwt"
        sign_req = urllib.request.Request(
            sign_url,
            data=json.dumps({"payload": json.dumps(jwt_payload)}).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {gcp_token}",
                "Content-Type": "application/json"
            },
            method="POST"
        )

        with urllib.request.urlopen(sign_req, timeout=5) as sign_resp:
            sign_data = json.loads(sign_resp.read().decode("utf-8"))
            signed_jwt = sign_data.get("signedJwt")

        if not signed_jwt:
            return None

        # Exchange signed JWT assertion with Google OAuth2 endpoint
        token_url = "https://oauth2.googleapis.com/token"
        token_data = urllib.parse.urlencode({
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "assertion": signed_jwt
        }).encode("utf-8")

        token_req = urllib.request.Request(
            token_url,
            data=token_data,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST"
        )

        with urllib.request.urlopen(token_req, timeout=5) as token_resp:
            oauth_data = json.loads(token_resp.read().decode("utf-8"))
            access_token = oauth_data.get("access_token")
            if access_token:
                _user_token_cache[user_email] = (access_token, now + 3500)
                logger.info(
                    f"[AUTO_DWD] Successfully minted automated Google OAuth access token for user '{user_email}'",
                    extra={"user_id": user_email, "token_type": "DWD_USER_TOKEN"}
                )
                return access_token

    except Exception as e:
        logger.warning(
            f"[AUTO_DWD_INFO] Automated user token generation for '{user_email}' (pending DWD admin authorization): {e}",
            extra={"user_id": user_email, "error": str(e)}
        )
        return None


def resolve_end_user_google_token(
    user: "AuthenticatedUser", 
    project_id: str, 
    location: str = "global"
) -> Tuple[Optional[str], str, Dict[str, Any]]:
    """
    Dynamically resolves the end-user Google token based on project Gemini Enterprise settings.
    Returns: (resolved_token, auth_mode, diagnostic_metadata)
    """
    metadata: Dict[str, Any] = {"project_id": project_id, "location": location}

    # 1. Determine Identity Mode (Auto-inspect aclConfig or use manual override)
    auth_mode = USER_AUTH_MODE
    if auth_mode == "auto":
        acl_config = get_discovery_engine_acl_config(project_id, location)
        idp_type = acl_config.get("idpConfig", {}).get("idpType", "GSUITE")
        metadata["discovered_idp_type"] = idp_type

        if idp_type == "THIRD_PARTY":
            auth_mode = "wif"
            workforce_pool = acl_config.get("idpConfig", {}).get("externalIdpConfig", {}).get("workforcePoolName", "")
            metadata["workforce_pool"] = workforce_pool
        else:
            auth_mode = "cloud_identity"
    
    metadata["effective_auth_mode"] = auth_mode

    # 2. Execute Token Resolution based on determined mode
    if auth_mode == "wif":
        workforce_pool = metadata.get("workforce_pool") or "locations/global/workforcePools/ca-entra-id-oidc-pool"
        if user.raw_token:
            wif_token = exchange_entra_jwt_for_wif_token(user.raw_token, workforce_pool)
            if wif_token:
                metadata["token_resolution_status"] = "WIF_TOKEN_EXCHANGED"
                return wif_token, "wif", metadata
        metadata["token_resolution_status"] = "WIF_EXCHANGE_SKIPPED_NO_RAW_TOKEN"
        return None, "wif", metadata

    elif auth_mode == "cloud_identity":
        cached_oauth3_tok, _ = get_cached_oauth3_google_token(user)
        if cached_oauth3_tok:
            metadata["token_resolution_status"] = "CACHED_OAUTH3_USER_TOKEN"
            return cached_oauth3_tok, "cloud_identity", metadata

        # Automatically attempt on-the-fly Google User Token generation for this Cloud Identity / Workspace user
        user_email = user.email or user.user_id
        dwd_token = mint_user_google_token_via_dwd(user_email)
        if dwd_token:
            metadata["token_resolution_status"] = "DWD_USER_TOKEN_MINTED"
            return dwd_token, "cloud_identity", metadata

        metadata["token_resolution_status"] = "CLOUD_IDENTITY_USER_ATTRIBUTED"
        return None, "cloud_identity", metadata

    return None, "service_account", metadata


def get_google_id_token(audience: str) -> Optional[str]:
    """
    Obtains a Google Cloud OIDC ID token to authenticate against downstream Cloud Run services.
    Attempts:
    1. Cloud Run / GCE Metadata server (Standard in GCP execution environment)
    2. google.oauth2.id_token (via Application Default Credentials / Service Account)
    """
    if not audience:
        return None

    # 1. Try Metadata Server (Instantaneous on Cloud Run)
    metadata_url = f"http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/identity?audience={audience}"
    try:
        req = urllib.request.Request(metadata_url, headers={"Metadata-Flavor": "Google"})
        with urllib.request.urlopen(req, timeout=3) as resp:
            token = resp.read().decode("utf-8").strip()
            if token:
                return token
    except Exception as meta_err:
        logger.debug(f"Metadata server ID token fetch skipped: {meta_err}")

    # 2. Try Google Auth library
    try:
        from google.auth.transport.requests import Request as GoogleAuthRequest
        from google.oauth2 import id_token
        auth_req = GoogleAuthRequest()
        token = id_token.fetch_id_token(auth_req, audience)
        if token:
            return token
    except Exception as ga_err:
        logger.debug(f"Google auth library ID token fetch skipped: {ga_err}")

    return None

# Optional HTTP Bearer security scheme (auto_error=False allows custom error handling)
security = HTTPBearer(auto_error=False)


# ---------------------------------------------------------------------------
# User Context & Token Verification Models
# ---------------------------------------------------------------------------
class AuthenticatedUser(BaseModel):
    user_id: str

    email: Optional[str] = None
    name: Optional[str] = None
    tenant_id: Optional[str] = None
    oid: Optional[str] = None
    sub: Optional[str] = None
    roles: List[str] = []
    scopes: List[str] = []
    raw_claims: Dict[str, Any] = Field(default_factory=dict)
    raw_token: Optional[str] = None


class DiagnosticLogEntry(BaseModel):
    timestamp: Optional[str] = None
    level: str = "INFO"
    category: str = "DIAGNOSTICS"
    message: str
    details: Optional[Dict[str, Any]] = None


class DiagnosticLogBatchRequest(BaseModel):
    logs: List[DiagnosticLogEntry] = []
    client_context: Optional[Dict[str, Any]] = None


def extract_user_from_payload(payload: Dict[str, Any], raw_token: Optional[str] = None) -> AuthenticatedUser:
    """Extracts standardized user identity claims from a decoded Entra ID JWT payload."""
    email = payload.get("email") or payload.get("preferred_username") or payload.get("upn")
    user_id = email or payload.get("sub") or payload.get("oid") or "anonymous_authenticated_user"
    name = payload.get("name") or payload.get("preferred_username") or user_id
    tenant_id = payload.get("tid")
    oid = payload.get("oid")
    sub = payload.get("sub")
    
    # Scopes and roles
    scopes = payload.get("scp", "").split() if payload.get("scp") else []
    roles = payload.get("roles", [])

    return AuthenticatedUser(
        user_id=user_id,
        email=email,
        name=name,
        tenant_id=tenant_id,
        oid=oid,
        sub=sub,
        roles=roles,
        scopes=scopes,
        raw_claims=payload,
        raw_token=raw_token
    )


async def verify_entra_token(
    request: Request,
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(security)
) -> Optional[AuthenticatedUser]:
    """
    FastAPI dependency to verify Microsoft Entra ID SSO tokens.
    1. Extracts Bearer token from the Authorization header.
    2. Fetches matching signing key from Microsoft JWKS endpoint.
    3. Validates RS256 signature and expiration timestamp.
    4. Validates audience (matching ENTRA_APP_ID or Application ID URI).
    5. Returns populated AuthenticatedUser context.
    """
    if not ENTRA_APP_ID:
        if REQUIRE_ENTRA_AUTH:
            logger.error(
                "SSO authentication required but MICROSOFT_ENTRA_APP_ID is not configured!",
                extra={"error_code": "AUTH_CONFIG_MISSING", "require_entra_auth": True}
            )
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="SSO authentication is required, but MICROSOFT_ENTRA_APP_ID is not configured on the server."
            )
        dev_email = request.headers.get("x-end-user-email")
        dev_id = request.headers.get("x-end-user-id") or dev_email or "dev_local_user"
        dev_name = request.headers.get("x-end-user-name") or (dev_email.split("@")[0] if dev_email else "Local Developer (Auth Bypassed)")
        logger.warning(f"MICROSOFT_ENTRA_APP_ID not configured; using dev user identity '{dev_id}'.")
        return AuthenticatedUser(user_id=dev_id, email=dev_email, name=dev_name)

    if not credentials or not credentials.credentials:
        if REQUIRE_ENTRA_AUTH:
            logger.warning(
                "Authentication rejected: Missing or empty Authorization Bearer token.",
                extra={"error_code": "AUTH_HEADER_MISSING", "status_code": 401}
            )
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Missing Authorization header with Bearer token.",
                headers={"WWW-Authenticate": "Bearer"}
            )
        dev_email = request.headers.get("x-end-user-email")
        dev_id = request.headers.get("x-end-user-id") or dev_email or "anonymous_dev_user"
        dev_name = request.headers.get("x-end-user-name") or (dev_email.split("@")[0] if dev_email else "Anonymous User (Auth Not Enforced)")
        logger.info(f"No Authorization token provided; using dev identity '{dev_id}' (REQUIRE_ENTRA_AUTH=false).")
        return AuthenticatedUser(user_id=dev_id, email=dev_email, name=dev_name)

    token = credentials.credentials
    token_len = len(token)

    # Inspect unverified header for diagnostics
    try:
        unverified_header = jwt.get_unverified_header(token)
        token_kid = unverified_header.get("kid")
        token_alg = unverified_header.get("alg")
    except Exception as e:
        logger.warning(
            f"Failed to parse unverified JWT header: {e}",
            extra={"error_code": "JWT_HEADER_PARSE_FAILED", "token_length": token_len}
        )
        unverified_header = {}
        token_kid = None
        token_alg = None

    if VERBOSE_LOGGING:
        logger.debug(
            f"Verifying incoming JWT token (kid: {token_kid}, alg: {token_alg}, length: {token_len})",
            extra={
                "token_header": unverified_header,
                "token_kid": token_kid,
                "token_alg": token_alg,
                "token_length": token_len
            }
        )

    try:
        # Fetch signing key matching token's kid from Microsoft JWKS
        signing_key = jwks_client.get_signing_key_from_jwt(token)
        
        # Decode and validate signature + expiration
        payload = jwt.decode(
            token,
            signing_key.key,
            algorithms=["RS256"],
            options={
                "verify_aud": False,  # Manually verified below for URI/GUID flexibility
                "verify_exp": True,
                "verify_iss": False,  # Verified explicitly below for tenant-specific or multi-tenant trust
            }
        )

        # Audience validation: Check if aud matches client GUID or ends with any configured app ID
        aud = payload.get("aud")
        if not aud:
            logger.error(
                "Token verification failed: Missing 'aud' claim in JWT payload.",
                extra={"error_code": "AUDIENCE_CLAIM_MISSING", "token_kid": token_kid}
            )
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Token is missing audience claim.")

        allowed_app_ids = [aid.strip() for aid in ENTRA_APP_ID.split(",") if aid.strip()]

        aud_valid = False
        for app_id in allowed_app_ids:
            if (
                aud == app_id or 
                (isinstance(aud, str) and (
                    aud == f"api://{app_id}"
                ))
            ):
                aud_valid = True
                break

        if not aud_valid:
            logger.error(
                f"Token audience mismatch. Expected client ID or URI matching one of {allowed_app_ids}, received '{aud}'",
                extra={
                    "error_code": "AUDIENCE_MISMATCH",
                    "expected_app_ids": allowed_app_ids,
                    "received_aud": aud,
                    "token_kid": token_kid
                }
            )
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail=f"Token audience mismatch. Expected client ID or URI matching one of {allowed_app_ids}."
            )

        # Strict tenant enforcement and issuer validation
        allowed_tenants = [t.strip() for t in ENTRA_TENANT_ID.split(",") if t.strip()]
        if not allowed_tenants and REQUIRE_ENTRA_AUTH:
            logger.error(
                "SSO authentication required but MICROSOFT_ENTRA_TENANT_ID is not configured!",
                extra={"error_code": "TENANT_CONFIG_MISSING", "require_entra_auth": True}
            )
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="SSO authentication is required, but MICROSOFT_ENTRA_TENANT_ID is not configured on the server."
            )

        iss = payload.get("iss", "")
        tid = payload.get("tid")
        if not tid:
            logger.error(
                "Token verification failed: Missing tenant ID ('tid') in token payload.",
                extra={"error_code": "TENANT_ID_MISSING", "token_kid": token_kid}
            )
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Token is missing tenant identifier."
            )

        if allowed_tenants:
            if tid not in allowed_tenants:
                logger.error(
                    f"Token verification failed: Tenant ID '{tid}' is not authorized. Allowed tenants: {allowed_tenants}",
                    extra={"error_code": "AUTH_INVALID_TENANT", "actual_tid": tid, "allowed_tenants": allowed_tenants}
                )
                raise HTTPException(
                    status_code=status.HTTP_401_UNAUTHORIZED,
                    detail="Token issued by untrusted tenant."
                )

            valid_issuers = []
            for tenant in allowed_tenants:
                valid_issuers.extend([
                    f"https://login.microsoftonline.com/{tenant}/v2.0",
                    f"https://sts.windows.net/{tenant}/"
                ])
        else:
            # Isolated dev mode when REQUIRE_ENTRA_AUTH is False
            valid_issuers = [
                f"https://login.microsoftonline.com/{tid}/v2.0",
                f"https://sts.windows.net/{tid}/"
            ]

        if not iss or iss not in valid_issuers:
            logger.error(
                f"Token issuer mismatch: received '{iss}', expected one of {valid_issuers}",
                extra={"error_code": "ISSUER_MISMATCH", "iss": iss, "valid_issuers": valid_issuers}
            )
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Token issuer mismatch or invalid."
            )

        user = extract_user_from_payload(payload, raw_token=token)

        # Cross-tenant enforcement on extracted user model
        if allowed_tenants and (not user.tenant_id or user.tenant_id not in allowed_tenants):
            logger.error(
                f"Token verification failed: User tenant ID '{user.tenant_id}' is not in allowed tenants: {allowed_tenants}",
                extra={"error_code": "AUTH_INVALID_TENANT", "actual_tid": user.tenant_id, "allowed_tenants": allowed_tenants}
            )
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Token issued by untrusted tenant."
            )

        # Default standard logging
        logger.info(
            f"Successfully authenticated Entra ID user: '{user.user_id}' (Tenant: {user.tenant_id or 'N/A'})",
            extra={
                "auth_status": "SUCCESS",
                "user_id": user.user_id,
                "tenant_id": user.tenant_id,
                "scopes": user.scopes,
                "roles": user.roles
            }
        )

        # Verbose deep diagnostic logging (All user parameters and decoded claims)
        if VERBOSE_LOGGING:
            user_dump = {
                "authenticated": True,
                "user_id": user.user_id,
                "name": user.name,
                "email": user.email,
                "tenant_id": user.tenant_id,
                "oid": user.oid,
                "sub": user.sub,
                "roles": user.roles,
                "scopes": user.scopes,
                "raw_claims": user.raw_claims
            }
            logger.debug(
                f"[VERBOSE AUTH CONTEXT] Authenticated User Claims: {json.dumps(user_dump)}",
                extra={"user_context": user_dump}
            )

        return user

    except jwt.ExpiredSignatureError:
        logger.warning(
            "Entra ID JWT token has expired.",
            extra={"error_code": "TOKEN_EXPIRED", "token_kid": token_kid}
        )
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token has expired. Please refresh the Office 365 access token.",
            headers={"WWW-Authenticate": "Bearer error=\"invalid_token\", error_description=\"The token has expired\""}
        )
    except jwt.InvalidTokenError as e:
        logger.warning(
            f"Invalid Entra ID JWT token: {str(e)}",
            extra={"error_code": "INVALID_TOKEN", "error_detail": str(e), "token_kid": token_kid}
        )
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Invalid token: {str(e)}",
            headers={"WWW-Authenticate": "Bearer error=\"invalid_token\""}
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(
            f"Unexpected token verification error: {str(e)}",
            extra={"error_code": "AUTH_UNEXPECTED_ERROR", "error_detail": str(e), "token_kid": token_kid},
            exc_info=True
        )
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Authentication error: {str(e)}"
        )


# ---------------------------------------------------------------------------
# FastAPI Application & Middleware Setup
# ---------------------------------------------------------------------------
app = FastAPI(
    title="GE Office 365 Assistant Auth-Proxy",
    description="Decoupled backend proxy handling Microsoft Entra ID JWT authentication for GE Office 365 Assistant Add-On (PowerPoint, Excel, Word).",
    version="1.0.0"
)

# Enable CORS for Microsoft Office webviews, Office Online, and local testing.
#
# No deployment URLs are hardcoded so the same build can run in parallel across
# many GCP projects and regions. The Cloud Run pattern accepts every URL format
# Cloud Run issues, in any project or region:
#   https://<service>-<project-number>.<region>.run.app   (deterministic, regional)
#   https://<service>-<hash>-<region-code>.a.run.app      (legacy hashed)
#   https://<service>.run.app
# Additional exact origins (e.g. a custom domain mapped to the frontend) can be
# supplied per deployment via CORS_ALLOWED_ORIGINS="https://a.example.com,https://b.example.com".
_CORS_BASE_REGEX = (
    r"^https?://(localhost|127\.0\.0\.1)(:[0-9]+)?$"
    r"|^https://([a-zA-Z0-9-]+\.)*(office\.com|officeapps\.live\.com)$"
    r"|^https://[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)?\.run\.app$"
)
_CORS_EXTRA_ORIGINS = [
    o.strip().rstrip("/")
    for o in os.environ.get("CORS_ALLOWED_ORIGINS", "").split(",")
    if o.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=_CORS_EXTRA_ORIGINS,
    allow_origin_regex=_CORS_BASE_REGEX,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def request_logging_middleware(request: Request, call_next):
    """
    HTTP Middleware:
    1. Tracks execution latency.
    2. Logs incoming HTTP requests and completed responses with GCP structured fields.
    3. Adds standard Cache-Control headers to prevent caching sensitive responses.
    """
    start_time = time.time()
    client_ip = request.headers.get("x-forwarded-for", request.client.host if request.client else "unknown")
    user_agent = request.headers.get("user-agent", "unknown")
    method = request.method
    path = request.url.path

    if VERBOSE_LOGGING and path not in ("/health", "/"):
        logger.debug(
            f"Incoming HTTP {method} {path} from {client_ip}",
            extra={
                "httpRequest": {
                    "requestMethod": method,
                    "requestUrl": str(request.url),
                    "userAgent": user_agent,
                    "remoteIp": client_ip,
                }
            }
        )

    response = await call_next(request)
    latency_ms = round((time.time() - start_time) * 1000, 2)

    # Standard no-cache headers
    response.headers["Cache-Control"] = "private, no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"

    # Log completed requests (suppress repetitive health probe logs unless verbose)
    if path not in ("/health", "/") or VERBOSE_LOGGING:
        log_level = logging.INFO if response.status_code < 400 else logging.WARNING
        logger.log(
            log_level,
            f"HTTP {method} {path} -> {response.status_code} ({latency_ms}ms)",
            extra={
                "httpRequest": {
                    "requestMethod": method,
                    "requestUrl": str(request.url),
                    "status": response.status_code,
                    "userAgent": user_agent,
                    "remoteIp": client_ip,
                    "latency": f"{latency_ms / 1000:.4f}s",
                },
                "latency_ms": latency_ms,
                "status_code": response.status_code
            }
        )

    return response


# ---------------------------------------------------------------------------
# Request / Response Schemas
# ---------------------------------------------------------------------------
class ChatMessage(BaseModel):
    role: str
    content: Optional[str] = None
    parts: Optional[List[Dict[str, Any]]] = None


class OfficeAddinRequest(BaseModel):
    prompt: str = Field(..., description="The user prompt or command")
    history: Optional[List[ChatMessage]] = Field(default=[], description="Chat history")
    sessionId: Optional[str] = Field(None, description="Active session ID")
    enableGrounding: Optional[bool] = Field(True, description="Enable grounding flag")
    userPseudoId: Optional[str] = Field(None, description="Optional client-provided pseudo user ID")
    attachments: Optional[List[Dict[str, Any]]] = Field(default=None, description="Optional document attachments for streamAssist")
    host: Optional[str] = Field(None, description="Active Office host surface (Word, Excel, PowerPoint, Outlook)")
    dataStoreId: Optional[str] = Field(None, description="Optional Data Store ID")


class HealthResponse(BaseModel):
    status: str
    service: str
    version: str
    require_entra_auth: bool
    entra_app_id_configured: bool
    verbose_logging: bool
    downstream_backend_configured: bool
    downstream_url: Optional[str] = None


# ---------------------------------------------------------------------------
# API Routes
# ---------------------------------------------------------------------------
@app.get("/", response_model=HealthResponse)
@app.get("/health", response_model=HealthResponse)
async def health_check():
    """Health check endpoint for Cloud Run and container probes."""
    return HealthResponse(
        status="healthy",
        service="auth-proxy",
        version="1.0.0",
        require_entra_auth=REQUIRE_ENTRA_AUTH,
        entra_app_id_configured=bool(ENTRA_APP_ID),
        verbose_logging=VERBOSE_LOGGING,
        downstream_backend_configured=bool(DOWNSTREAM_BACKEND_URL),
        downstream_url=DOWNSTREAM_BACKEND_URL or None
    )


@app.get("/api/config")
async def get_public_config():
    """
    Returns public frontend configuration dynamically so no Client IDs
    or environment-specific parameters are hardcoded into the frontend code.
    """
    return {
        "google_oauth_client_id": GOOGLE_OAUTH_CLIENT_ID,
        "microsoft_entra_app_id": ENTRA_APP_ID,
        "user_auth_mode": USER_AUTH_MODE,
        "gcp_project_id": GCP_PROJECT_ID,
        "require_entra_auth": REQUIRE_ENTRA_AUTH
    }


class GoogleTokenSyncRequest(BaseModel):
    google_token: Optional[str] = None
    expires_in: Optional[int] = 3600


@app.get("/api/auth/me")
async def get_current_user(user: AuthenticatedUser = Depends(verify_entra_token)):
    """
    Returns the authenticated user details extracted from the Entra ID JWT token,
    including any active 3-legged Google OAuth token cached across Office surfaces.
    """
    cached_google_token, remaining_sec = get_cached_oauth3_google_token(user)

    if not cached_google_token and DOWNSTREAM_BACKEND_URL:
        try:
            google_id_token = get_google_id_token(DOWNSTREAM_BACKEND_URL)
            headers = {
                "Accept": "application/json",
                "X-End-User-Id": user.user_id,
                "X-End-User-Email": user.email or user.user_id,
                "X-End-User-Oid": user.oid or "",
                "X-End-User-Tenant-Id": user.tenant_id or "",
            }
            if google_id_token:
                headers["Authorization"] = f"Bearer {google_id_token}"
            async with httpx.AsyncClient(timeout=5.0) as client:
                resp = await client.get(f"{DOWNSTREAM_BACKEND_URL}/api/auth/me", headers=headers)
                if resp.status_code == 200:
                    d_data = resp.json()
                    if d_data.get("google_token"):
                        cached_google_token = d_data["google_token"]
                        remaining_sec = int(d_data.get("google_token_expires_in") or 1800)
                        cache_oauth3_google_token(user, cached_google_token, remaining_sec)
        except Exception as e:
            logger.debug(f"Downstream /api/auth/me lookup skipped: {e}")

    user_response = {
        "authenticated": True,
        "user_id": user.user_id,
        "name": user.name,
        "email": user.email,
        "tenant_id": user.tenant_id,
        "oid": user.oid,
        "sub": user.sub,
        "roles": user.roles,
        "scopes": user.scopes,
        "google_token": cached_google_token,
        "google_token_expires_in": remaining_sec if cached_google_token else 0,
    }
    if VERBOSE_LOGGING:
        logger.debug(
            f"Returning /api/auth/me for user: {user.user_id}",
            extra={"user_profile": {k: v for k, v in user_response.items() if k != "google_token"}}
        )
    return user_response


@app.post("/api/auth/google-token")
async def sync_user_google_token(
    payload: GoogleTokenSyncRequest,
    user: AuthenticatedUser = Depends(verify_entra_token)
):
    """
    Stores or clears the authenticated Entra ID user's 3-legged Google OAuth token
    so all 4 Office surfaces (Outlook, Excel, Word, PowerPoint) share the login session.
    """
    cache_oauth3_google_token(user, payload.google_token, payload.expires_in or 3600)

    if DOWNSTREAM_BACKEND_URL:
        try:
            google_id_token = get_google_id_token(DOWNSTREAM_BACKEND_URL)
            headers = {
                "Content-Type": "application/json",
                "X-End-User-Id": user.user_id,
                "X-End-User-Email": user.email or user.user_id,
                "X-End-User-Oid": user.oid or "",
                "X-End-User-Tenant-Id": user.tenant_id or "",
            }
            if google_id_token:
                headers["Authorization"] = f"Bearer {google_id_token}"
            async with httpx.AsyncClient(timeout=5.0) as client:
                await client.post(
                    f"{DOWNSTREAM_BACKEND_URL}/api/auth/google-token",
                    json={"google_token": payload.google_token, "expires_in": payload.expires_in or 3600},
                    headers=headers
                )
        except Exception as e:
            logger.debug(f"Downstream /api/auth/google-token sync skipped: {e}")

    return {
        "status": "ok",
        "connected": bool(payload.google_token and payload.google_token.strip()),
        "user_id": user.user_id
    }


@app.delete("/api/auth/google-token")
async def clear_user_google_token(user: AuthenticatedUser = Depends(verify_entra_token)):
    """Clears the cached 3-legged Google OAuth token for the authenticated Entra ID user."""
    cache_oauth3_google_token(user, None)
    if DOWNSTREAM_BACKEND_URL:
        try:
            google_id_token = get_google_id_token(DOWNSTREAM_BACKEND_URL)
            headers = {
                "Content-Type": "application/json",
                "X-End-User-Id": user.user_id,
                "X-End-User-Email": user.email or user.user_id,
                "X-End-User-Oid": user.oid or "",
                "X-End-User-Tenant-Id": user.tenant_id or "",
            }
            if google_id_token:
                headers["Authorization"] = f"Bearer {google_id_token}"
            async with httpx.AsyncClient(timeout=5.0) as client:
                await client.post(
                    f"{DOWNSTREAM_BACKEND_URL}/api/auth/google-token",
                    json={"google_token": None, "expires_in": 0},
                    headers=headers
                )
        except Exception as e:
            logger.debug(f"Downstream /api/auth/google-token clear skipped: {e}")
    return {"status": "ok", "connected": False, "user_id": user.user_id}


@app.post("/api/auth/validate")
async def validate_token_endpoint(user: AuthenticatedUser = Depends(verify_entra_token)):
    """
    Explicit endpoint to validate an incoming Bearer token.
    Returns 200 OK with decoded claims if valid; raises 401 if invalid.
    """
    return {
        "valid": True,
        "message": "Entra ID token is valid and active.",
        "user": {
            "user_id": user.user_id,
            "name": user.name,
            "email": user.email,
            "tenant_id": user.tenant_id
        }
    }


@app.post("/api/diagnostics/log")
async def ingest_client_diagnostics(
    req: DiagnosticLogBatchRequest,
    request: Request
):
    """
    Ingests client-side troubleshooting and diagnostic logs from the Office 365 Add-in
    and writes structured JSON records directly to Google Cloud Logging.
    """
    client_ctx = req.client_context or {}
    client_ip = request.client.host if request.client else "unknown"
    client_ctx["client_ip"] = client_ip

    for entry in req.logs:
        lvl = (entry.level or "INFO").upper()
        cat = entry.category or "DIAGNOSTICS"
        structured_payload = {
            "logger": "client-diagnostics",
            "category": cat,
            "client_timestamp": entry.timestamp,
            "client_context": client_ctx,
            "details": entry.details or {},
            "user_id": client_ctx.get("user_email") or client_ctx.get("user_id") or "anonymous"
        }
        
        log_msg = f"[CLIENT_DIAGNOSTICS][{cat}] {entry.message}"
        
        if lvl == "ERROR":
            logger.error(log_msg, extra=structured_payload)
        elif lvl == "WARNING":
            logger.warning(log_msg, extra=structured_payload)
        elif lvl == "DEBUG":
            logger.debug(log_msg, extra=structured_payload)
        else:
            logger.info(log_msg, extra=structured_payload)

    return {"status": "ok", "ingested": len(req.logs)}


@app.post("/askGeminiEnterprise")
@app.post("/api/gemini/chat")
async def proxy_addin_request(
    req: OfficeAddinRequest, 
    request: Request,
    user: AuthenticatedUser = Depends(verify_entra_token)
):
    """
    Decoupled endpoint for Office 365 add-ins (PowerPoint, Excel, Word).
    1. Validates Entra ID JWT authentication and captures end-user identity.
    2. If DOWNSTREAM_BACKEND_URL is configured:
       - Obtains a Google Cloud S2S IAM token for the downstream service.
       - Forwards prompt, sessionId, and user identity headers to geminiproxy.
       - Returns the grounded response with authenticated user context.
    3. If DOWNSTREAM_BACKEND_URL is not configured:
       - Returns a structured verification response (isolated dev mode).
    """
    if not req.prompt:
        logger.warning(
            "Rejected request to /askGeminiEnterprise: Prompt is empty.",
            extra={"user_id": user.user_id}
        )
        raise HTTPException(status_code=400, detail="Prompt is required.")

    # Captured end-user ID from verified JWT
    authenticated_user_id = user.user_id
    session_id = req.sessionId if req.sessionId else None

    logger.info(
        f"Processing add-in request from '{authenticated_user_id}' (Session: {session_id or 'NEW'}, Prompt length: {len(req.prompt)} chars)",
        extra={
            "user_id": authenticated_user_id,
            "sessionId": session_id or "NEW",
            "prompt_preview": req.prompt[:100] + ("..." if len(req.prompt) > 100 else ""),
            "enableGrounding": req.enableGrounding
        }
    )

    if VERBOSE_LOGGING:
        logger.debug(
            f"[VERBOSE PAYLOAD] Full prompt received from '{authenticated_user_id}': {req.prompt}",
            extra={
                "user_id": authenticated_user_id,
                "full_prompt": req.prompt,
                "history_count": len(req.history) if req.history else 0,
                "sessionId": session_id,
                "user_parameters": {
                    "authenticated": True,
                    "user_id": user.user_id,
                    "name": user.name,
                    "email": user.email,
                    "tenant_id": user.tenant_id,
                    "oid": user.oid,
                    "sub": user.sub,
                    "roles": user.roles,
                    "scopes": user.scopes
                }
            }
        )

    # If downstream backend is configured, proxy request with Google S2S IAM authentication
    if DOWNSTREAM_BACKEND_URL:
        target_endpoint = f"{DOWNSTREAM_BACKEND_URL}/askGeminiEnterprise"
        logger.info(
            f"Forwarding request to downstream backend: {target_endpoint}",
            extra={"downstream_endpoint": target_endpoint, "user_id": authenticated_user_id}
        )

        # 1. Check if 3-legged user OAuth token was passed directly from the Office 365 add-in
        passed_google_token = request.headers.get("x-end-user-google-token")
        if passed_google_token:
            cache_oauth3_google_token(user, passed_google_token)
            user_google_token = passed_google_token
            detected_auth_mode = USER_AUTH_MODE if USER_AUTH_MODE != "auto" else "cloud_identity"
            auth_diag = {
                "project_id": GCP_PROJECT_ID,
                "location": GCP_LOCATION,
                "effective_auth_mode": detected_auth_mode,
                "token_resolution_status": "PASSED_VIA_HEADER_OAUTH3"
            }
        else:
            # Dynamically resolve end-user Google token based on project Gemini Enterprise settings (DWD / WIF)
            user_google_token, detected_auth_mode, auth_diag = resolve_end_user_google_token(
                user=user,
                project_id=GCP_PROJECT_ID,
                location=GCP_LOCATION
            )

        effective_user_id = authenticated_user_id
        effective_email = user.email or authenticated_user_id
        effective_name = user.name or authenticated_user_id

        headers = {
            "Content-Type": "application/json",
            "X-End-User-Id": effective_user_id,
            "X-End-User-Email": effective_email,
            "X-End-User-Name": effective_name,
            "X-End-User-Tenant-Id": user.tenant_id or "",
            "X-User-Auth-Mode": detected_auth_mode
        }

        if user_google_token:
            headers["X-End-User-Google-Token"] = user_google_token
            if VERBOSE_LOGGING:
                logger.debug(
                    f"Attached resolved end-user Google token for user '{effective_user_id}' (Mode: {detected_auth_mode})",
                    extra={"auth_mode": detected_auth_mode, "user_id": effective_user_id}
                )
        else:
            logger.info(
                f"No end-user Google token attached (Mode: {detected_auth_mode}, Status: {auth_diag.get('token_resolution_status')})",
                extra={"auth_diag": auth_diag, "user_id": effective_user_id}
            )

        google_id_token = get_google_id_token(DOWNSTREAM_BACKEND_URL)
        if google_id_token:
            headers["Authorization"] = f"Bearer {google_id_token}"
            if VERBOSE_LOGGING:
                logger.debug("Attached Google IAM ID token for downstream S2S invocation.")
        else:
            logger.warning("Could not obtain Google IAM ID token; invoking downstream service.")

        effective_host = (req.host or request.headers.get("x-office-host", "")).strip() or None
        if effective_host:
            headers["X-Office-Host"] = effective_host

        downstream_payload = {
            "prompt": req.prompt,
            "history": [msg.dict() for msg in req.history] if req.history else [],
            "sessionId": session_id,
            "enableGrounding": req.enableGrounding,
            "host": effective_host,
            "dataStoreId": req.dataStoreId,
            "userPseudoId": effective_user_id,
            "userId": effective_user_id,
            "authenticatedUser": {
                "userId": effective_user_id,
                "email": effective_email,
                "name": effective_name,
                "tenantId": user.tenant_id,
                "oid": user.oid,
                "sub": user.sub
            }
        }
        if req.attachments:
            downstream_payload["attachments"] = req.attachments
            logger.info(
                f"Forwarding {len(req.attachments)} document attachment(s) downstream to streamAssist",
                extra={"user_id": effective_user_id, "attachment_count": len(req.attachments)}
            )

        try:
            downstream_timeout = int(os.environ.get("DOWNSTREAM_TIMEOUT", "300"))
            resp = requests.post(
                target_endpoint,
                json=downstream_payload,
                headers=headers,
                timeout=downstream_timeout
            )

            if resp.status_code >= 400:
                logger.error(
                    f"Downstream service returned error HTTP {resp.status_code}: {resp.text}",
                    extra={
                        "downstream_status": resp.status_code,
                        "downstream_endpoint": target_endpoint,
                        "error_detail": resp.text[:500]
                    }
                )
                raise HTTPException(
                    status_code=resp.status_code,
                    detail=f"Downstream backend returned error {resp.status_code}: {resp.text}"
                )

            data = resp.json() if resp.headers.get("content-type", "").startswith("application/json") else {"result": resp.text}
            if isinstance(data, dict) and "authenticatedUser" not in data:
                data["authenticatedUser"] = {
                    "userId": authenticated_user_id,
                    "email": user.email,
                    "name": user.name,
                    "tenantId": user.tenant_id,
                }
            return data

        except requests.exceptions.RequestException as req_err:
            logger.error(
                f"Failed to communicate with downstream Gemini backend: {str(req_err)}",
                extra={"error_code": "DOWNSTREAM_COMMUNICATION_ERROR", "error_detail": str(req_err)},
                exc_info=True
            )
            raise HTTPException(
                status_code=502,
                detail=f"Failed to communicate with downstream Gemini backend: {str(req_err)}"
            )

    # In isolated dev mode (when DOWNSTREAM_BACKEND_URL is not set), return structured auth confirmation
    return {
        "result": f"[Auth-Proxy Verified] Hello {user.name or authenticated_user_id}! Your request was authenticated via Microsoft Entra ID. User ID captured: '{authenticated_user_id}'. Prompt received: '{req.prompt}'.",
        "sessionId": session_id,
        "authenticatedUser": {
            "userId": authenticated_user_id,
            "email": user.email,
            "name": user.name,
            "tenantId": user.tenant_id,
        },
        "citations": [],
        "backendMode": "auth-proxy-isolated"
    }


@app.post("/askGeminiStream")
async def ask_gemini_stream(
    req: OfficeAddinRequest,
    request: Request,
    user: AuthenticatedUser = Depends(verify_entra_token)
):
    """
    Authenticated SSE streaming endpoint for Gemini Enterprise (used by Outlook add-in).
    Validates Microsoft Entra ID token, resolves end-user Google token (3LO/WIF/DWD),
    mints S2S Google OIDC ID token, and pipes SSE chunks from downstream /askGeminiStream.
    """
    if not req.prompt or not req.prompt.strip():
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Missing 'prompt' field in request payload."
        )

    authenticated_user_id = user.email or user.user_id
    session_id = req.sessionId if req.sessionId else None

    passed_google_token = request.headers.get("x-end-user-google-token", "").strip()
    if passed_google_token:
        cache_oauth3_google_token(user, passed_google_token)
        user_google_token = passed_google_token
        detected_auth_mode = USER_AUTH_MODE if USER_AUTH_MODE != "auto" else "cloud_identity"
    else:
        user_google_token, detected_auth_mode, _ = resolve_end_user_google_token(
            user=user,
            project_id=GCP_PROJECT_ID,
            location=GCP_LOCATION
        )

    if DOWNSTREAM_BACKEND_URL:
        google_id_token = get_google_id_token(DOWNSTREAM_BACKEND_URL)
        effective_user_id = authenticated_user_id
        effective_email = user.email or authenticated_user_id
        effective_name = user.name or authenticated_user_id

        downstream_headers = {
            "Content-Type": "application/json",
            "Accept": "text/event-stream",
            "X-End-User-Id": effective_user_id,
            "X-End-User-Email": effective_email,
            "X-End-User-Name": effective_name,
            "X-End-User-Tenant-Id": user.tenant_id or "",
            "X-User-Auth-Mode": detected_auth_mode,
        }
        effective_host = (req.host or request.headers.get("x-office-host", "")).strip() or None
        if google_id_token:
            downstream_headers["Authorization"] = f"Bearer {google_id_token}"
        if user_google_token:
            downstream_headers["X-End-User-Google-Token"] = user_google_token
        if effective_host:
            downstream_headers["X-Office-Host"] = effective_host

        downstream_payload = {
            "prompt": req.prompt,
            "history": [msg.dict(exclude_none=True) for msg in req.history] if req.history else [],
            "sessionId": session_id,
            "enableGrounding": req.enableGrounding,
            "host": effective_host,
            "dataStoreId": req.dataStoreId,
            "userPseudoId": effective_user_id,
            "userId": effective_user_id,
            "authenticatedUser": {
                "userId": effective_user_id,
                "email": effective_email,
                "name": effective_name,
                "tenantId": user.tenant_id,
                "oid": user.oid,
                "sub": user.sub,
            },
        }
        if req.attachments:
            downstream_payload["attachments"] = req.attachments

        target_endpoint = f"{DOWNSTREAM_BACKEND_URL}/askGeminiStream"
        downstream_timeout = int(os.environ.get("DOWNSTREAM_TIMEOUT", "180"))

        async def event_generator():
            async with httpx.AsyncClient(timeout=downstream_timeout) as client:
                async with client.stream(
                    "POST",
                    target_endpoint,
                    json=downstream_payload,
                    headers=downstream_headers
                ) as resp:
                    if resp.status_code >= 400:
                        err_body = await resp.aread()
                        err_msg = err_body.decode("utf-8", errors="replace")[:500]
                        yield f"data: {json.dumps({'error': f'Downstream error ({resp.status_code}): {err_msg}', 'done': True})}\n\n"
                        return
                    async for chunk in resp.aiter_bytes():
                        if chunk:
                            yield chunk

        return StreamingResponse(
            event_generator(),
            media_type="text/event-stream",
            headers={
                "Cache-Control": "no-cache, no-transform",
                "Connection": "keep-alive",
                "X-Accel-Buffering": "no",
            }
        )

    # Isolated dev mode fallback
    async def dev_stream():
        msg = f"[Auth-Proxy Verified Stream] Hello {user.name or authenticated_user_id}! Prompt received: '{req.prompt}'."
        yield f"data: {json.dumps({'delta': msg, 'text': msg, 'done': False})}\n\n"
        yield f"data: {json.dumps({'delta': '', 'text': '', 'fullText': msg, 'sessionId': session_id, 'citations': [], 'done': True})}\n\n"
        yield "data: [DONE]\n\n"

    return StreamingResponse(dev_stream(), media_type="text/event-stream")


class DeleteSessionPayload(BaseModel):
    sessionId: str


@app.get("/sessions")
async def proxy_list_sessions(
    request: Request,
    pageSize: int = 20,
    host: Optional[str] = None,
    user: AuthenticatedUser = Depends(verify_entra_token)
):
    """
    Authenticated proxy endpoint to list enterprise sessions from downstream geminiproxy.
    """
    if DOWNSTREAM_BACKEND_URL:
        google_id_token = get_google_id_token(DOWNSTREAM_BACKEND_URL)
        effective_host = (host or request.headers.get("x-office-host", "")).strip() or None
        headers = {
            "Accept": "application/json",
            "X-End-User-Id": user.user_id,
            "X-End-User-Email": user.email or user.user_id,
            "X-End-User-Tenant-Id": user.tenant_id or "",
        }
        if effective_host:
            headers["X-Office-Host"] = effective_host
        user_google_token = request.headers.get("X-End-User-Google-Token", "").strip()
        if user_google_token:
            cache_oauth3_google_token(user, user_google_token)
        else:
            user_google_token, _ = get_cached_oauth3_google_token(user)
        if user_google_token:
            headers["X-End-User-Google-Token"] = user_google_token
            headers["X-User-Auth-Mode"] = "user_token"
        if google_id_token:
            headers["Authorization"] = f"Bearer {google_id_token}"

        target = f"{DOWNSTREAM_BACKEND_URL}/sessions?pageSize={min(max(pageSize, 1), 100)}"
        if effective_host:
            target += f"&host={effective_host}"
        async with httpx.AsyncClient(timeout=30.0) as client:
            resp = await client.get(target, headers=headers)
            if resp.status_code >= 400:
                raise HTTPException(status_code=resp.status_code, detail=resp.text)
            return resp.json()
    return {"sessions": []}


@app.post("/deleteSession")
async def proxy_delete_session(
    request: Request,
    payload: DeleteSessionPayload,
    user: AuthenticatedUser = Depends(verify_entra_token)
):
    """
    Authenticated proxy endpoint to delete an enterprise session from downstream geminiproxy.
    """
    if not payload.sessionId or not payload.sessionId.strip():
        raise HTTPException(status_code=400, detail="sessionId is required.")

    if DOWNSTREAM_BACKEND_URL:
        google_id_token = get_google_id_token(DOWNSTREAM_BACKEND_URL)
        headers = {
            "Content-Type": "application/json",
            "X-End-User-Id": user.user_id,
            "X-End-User-Email": user.email or user.user_id,
            "X-End-User-Tenant-Id": user.tenant_id or "",
        }
        user_google_token = request.headers.get("X-End-User-Google-Token", "").strip()
        if user_google_token:
            cache_oauth3_google_token(user, user_google_token)
        else:
            user_google_token, _ = get_cached_oauth3_google_token(user)
        if user_google_token:
            headers["X-End-User-Google-Token"] = user_google_token
            headers["X-User-Auth-Mode"] = "user_token"
        if google_id_token:
            headers["Authorization"] = f"Bearer {google_id_token}"

        target = f"{DOWNSTREAM_BACKEND_URL}/deleteSession"
        async with httpx.AsyncClient(timeout=30.0) as client:
            resp = await client.post(target, json={"sessionId": payload.sessionId}, headers=headers)
            if resp.status_code >= 400:
                raise HTTPException(status_code=resp.status_code, detail=resp.text)
            return resp.json()
    return {"success": True, "sessionId": payload.sessionId}


# ---------------------------------------------------------------------------
# CLI Direct Execution
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8080))
    logger.info(f"Starting auth-proxy server on port {port}...")
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=True)

