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

"""
Unit & Integration Tests for Auth-Proxy FastAPI Microservice
"""
import os
import unittest
from fastapi.testclient import TestClient

# Set test environment defaults before importing app
os.environ["REQUIRE_ENTRA_AUTH"] = "false"
os.environ["MICROSOFT_ENTRA_APP_ID"] = "test-client-guid-12345"
os.environ["MICROSOFT_ENTRA_TENANT_ID"] = "trusted-tenant-123"
os.environ["USER_AUTH_MODE"] = "auto"

from main import app, extract_user_from_payload, AuthenticatedUser

client = TestClient(app)


class TestAuthProxy(unittest.TestCase):

    def test_health_check(self):
        """Verify health check endpoints return 200 OK and expected structure."""
        response = client.get("/health")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["status"], "healthy")
        self.assertEqual(data["service"], "auth-proxy")
        self.assertTrue(data["entra_app_id_configured"])
        self.assertIn("verbose_logging", data)

    def test_root_endpoint(self):
        """Verify root / endpoint maps to health check."""
        response = client.get("/")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["status"], "healthy")
        self.assertIn("verbose_logging", data)

    def test_extract_user_from_payload(self):
        """Verify claim extraction helper properly captures user identity."""
        mock_payload = {
            "preferred_username": "john.doe@contoso.com",
            "name": "John Doe",
            "tid": "tenant-abc-123",
            "oid": "user-oid-456",
            "sub": "subject-789",
            "scp": "access_as_user profile email",
            "roles": ["Admin", "PowerUser"]
        }
        user = extract_user_from_payload(mock_payload)
        self.assertEqual(user.user_id, "john.doe@contoso.com")
        self.assertEqual(user.email, "john.doe@contoso.com")
        self.assertEqual(user.name, "John Doe")
        self.assertEqual(user.tenant_id, "tenant-abc-123")
        self.assertEqual(user.oid, "user-oid-456")
        self.assertIn("access_as_user", user.scopes)
        self.assertIn("Admin", user.roles)

    def test_extract_user_from_payload_with_raw_token(self):
        """Verify claim extraction helper properly captures user identity and raw token."""
        mock_payload = {
            "preferred_username": "alexw@5m4qby.onmicrosoft.com",
            "name": "Alex Wilber",
            "tid": "tenant-abc-123",
            "oid": "user-oid-456",
            "sub": "subject-789"
        }
        user = extract_user_from_payload(mock_payload, raw_token="mock_raw_jwt_token_123")
        self.assertEqual(user.user_id, "alexw@5m4qby.onmicrosoft.com")
        self.assertEqual(user.raw_token, "mock_raw_jwt_token_123")

    def test_resolve_end_user_google_token_gsuite(self):
        """Verify resolve_end_user_google_token correctly determines cloud_identity mode for GSUITE projects."""
        from main import resolve_end_user_google_token
        user = AuthenticatedUser(user_id="alexw@5m4qby.onmicrosoft.com", raw_token="mock_jwt")
        token, mode, diag = resolve_end_user_google_token(user, project_id="agentspace-452714", location="global")
        self.assertEqual(mode, "cloud_identity")
        self.assertIn("discovered_idp_type", diag)
        self.assertEqual(diag["discovered_idp_type"], "GSUITE")

    def test_resolve_end_user_google_token_wif(self):
        """Verify resolve_end_user_google_token correctly determines wif mode for THIRD_PARTY projects."""
        from main import resolve_end_user_google_token
        user = AuthenticatedUser(user_id="alexw@5m4qby.onmicrosoft.com", raw_token="mock_jwt")
        token, mode, diag = resolve_end_user_google_token(user, project_id="agentspace-wif", location="global")
        self.assertEqual(mode, "wif")
        self.assertIn("workforce_pool", diag)
        self.assertEqual(diag["workforce_pool"], "locations/global/workforcePools/ca-entra-id-oidc-pool")

    def test_entra_issuer_validation_logic(self):
        """Verify strict tenant issuer validation accepts genuine tenant issuers and rejects foreign/spoofed issuers."""
        import asyncio
        from unittest.mock import MagicMock, patch
        from fastapi import HTTPException
        from fastapi.security import HTTPAuthorizationCredentials
        from main import verify_entra_token

        mock_request = MagicMock()
        mock_creds = HTTPAuthorizationCredentials(scheme="Bearer", credentials="valid_format_jwt_token")

        with patch("main.jwks_client.get_signing_key_from_jwt") as mock_jwks, \
             patch("main.jwt.decode") as mock_decode, \
             patch("main.jwt.get_unverified_header") as mock_header:
            
            mock_header.return_value = {"kid": "key-1", "alg": "RS256"}
            mock_jwks.return_value.key = "fake_key"

            # 1. Foreign tenant issuer must be REJECTED even if starting with https://login.microsoftonline.com/
            mock_decode.return_value = {
                "aud": "test-client-guid-12345",
                "tid": "trusted-tenant-123",
                "iss": "https://login.microsoftonline.com/attacker-tenant-456/v2.0",
                "preferred_username": "attacker@evil.com",
                "sub": "sub-1"
            }
            with self.assertRaises(HTTPException) as ctx:
                asyncio.run(verify_entra_token(mock_request, mock_creds))
            self.assertEqual(ctx.exception.status_code, 401)
            self.assertIn("Token issuer mismatch", ctx.exception.detail)

            # 2. Spoofed domain prefix must be REJECTED
            mock_decode.return_value = {
                "aud": "test-client-guid-12345",
                "tid": "trusted-tenant-123",
                "iss": "https://login.microsoftonline.com.attacker.com/trusted-tenant-123/v2.0",
                "preferred_username": "attacker@evil.com",
                "sub": "sub-2"
            }
            with self.assertRaises(HTTPException) as ctx:
                asyncio.run(verify_entra_token(mock_request, mock_creds))
            self.assertEqual(ctx.exception.status_code, 401)

            # 3. Legitimate v2.0 matching tenant issuer must be ACCEPTED
            mock_decode.return_value = {
                "aud": "test-client-guid-12345",
                "tid": "trusted-tenant-123",
                "iss": "https://login.microsoftonline.com/trusted-tenant-123/v2.0",
                "preferred_username": "user@trusted.com",
                "name": "Trusted User",
                "sub": "sub-3"
            }
            user = asyncio.run(verify_entra_token(mock_request, mock_creds))
            self.assertEqual(user.user_id, "user@trusted.com")
            self.assertEqual(user.tenant_id, "trusted-tenant-123")

            # 4. Legitimate v1.0 sts.windows.net matching tenant issuer must be ACCEPTED
            mock_decode.return_value = {
                "aud": "test-client-guid-12345",
                "tid": "trusted-tenant-123",
                "iss": "https://sts.windows.net/trusted-tenant-123/",
                "preferred_username": "user2@trusted.com",
                "name": "Trusted User 2",
                "sub": "sub-4"
            }
            user2 = asyncio.run(verify_entra_token(mock_request, mock_creds))
            self.assertEqual(user2.user_id, "user2@trusted.com")

            # 5. Attacker presenting a token from a rogue tenant must be REJECTED (CWE-287 cross-tenant bypass prevention)
            mock_decode.return_value = {
                "aud": "test-client-guid-12345",
                "tid": "attacker-tenant-666",
                "iss": "https://login.microsoftonline.com/attacker-tenant-666/v2.0",
                "preferred_username": "attacker@evilcorp.com",
                "name": "Attacker",
                "sub": "sub-evil"
            }
            with patch("main.ENTRA_TENANT_ID", "trusted-tenant-123"):
                with self.assertRaises(HTTPException) as ctx:
                    asyncio.run(verify_entra_token(mock_request, mock_creds))
                self.assertEqual(ctx.exception.status_code, 401)
                self.assertIn("untrusted tenant", ctx.exception.detail)

            # 6. Token completely missing tenant identifier (tid) must be REJECTED
            mock_decode.return_value = {
                "aud": "test-client-guid-12345",
                "iss": "https://login.microsoftonline.com/trusted-tenant-123/v2.0",
                "preferred_username": "user@trusted.com",
                "sub": "sub-no-tid"
            }
            with self.assertRaises(HTTPException) as ctx:
                asyncio.run(verify_entra_token(mock_request, mock_creds))
            self.assertEqual(ctx.exception.status_code, 401)
            self.assertIn("tenant identifier", ctx.exception.detail)

            # 7. When REQUIRE_ENTRA_AUTH=True, missing ENTRA_TENANT_ID must fail safe with 500 error
            with patch("main.ENTRA_TENANT_ID", ""), patch("main.REQUIRE_ENTRA_AUTH", True):
                with self.assertRaises(HTTPException) as ctx:
                    asyncio.run(verify_entra_token(mock_request, mock_creds))
                self.assertEqual(ctx.exception.status_code, 500)
                self.assertIn("MICROSOFT_ENTRA_TENANT_ID is not configured", ctx.exception.detail)

    def test_session_proxy_endpoints(self):
        """Verify /sessions and /deleteSession proxy routes respond correctly in isolated mode."""
        # GET /sessions without downstream backend returns empty list in dev mode
        res_list = client.get("/sessions")
        self.assertEqual(res_list.status_code, 200)
        self.assertEqual(res_list.json(), {"sessions": []})

        # POST /deleteSession with valid body returns success in dev mode
        res_del = client.post("/deleteSession", json={"sessionId": "session-123"})
        self.assertEqual(res_del.status_code, 200)
        self.assertEqual(res_del.json(), {"success": True, "sessionId": "session-123"})

        # POST /deleteSession with empty sessionId returns 400
        res_del_bad = client.post("/deleteSession", json={"sessionId": ""})
        self.assertEqual(res_del_bad.status_code, 400)


if __name__ == "__main__":
    unittest.main()

