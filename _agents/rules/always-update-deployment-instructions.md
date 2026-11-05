# Rule: Keep Deployment Instructions in Sync

**Rule**: Whenever any changes are made to the deployment scripts (`deploy.sh`), environment configuration (`.env.example`), infrastructure architecture, or deployment flags, the `docs/deployment/cloud-run-deployment-guide.md` file MUST be updated in the same breath.

**Context**: The `docs/deployment/cloud-run-deployment-guide.md` is the source of truth for manual and automated enterprise deployments. If it falls out of sync with `deploy.sh` or the actual required gcloud commands (like IAM bindings or `--no-allow-unauthenticated` flags), users will experience deployment failures and security vulnerabilities.

**Checklist before ending a deployment-related task**:
1. Did I add a new environment variable? -> Update `.env.example` AND `docs/deployment/cloud-run-deployment-guide.md`.
2. Did I change a gcloud flag (e.g., `--allow-unauthenticated` to `--no-allow-unauthenticated`)? -> Find all occurrences in `docs/deployment/cloud-run-deployment-guide.md` and update them.
3. Did I add a new script (`deploy.sh`, `generate_manifest.py`)? -> Ensure `docs/deployment/cloud-run-deployment-guide.md` references it if applicable.
4. Did I add a new IAM policy requirement? -> Explicitly add the `gcloud run services add-iam-policy-binding` command to the instructions.
