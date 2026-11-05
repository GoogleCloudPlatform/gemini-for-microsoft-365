# How to contribute

We'd love to accept your patches and contributions to this project.

## Before you begin

### Sign our Contributor License Agreement

Contributions to this project must be accompanied by a
[Contributor License Agreement](https://cla.developers.google.com/about) (CLA).
You (or your employer) retain the copyright to your contribution; this simply
gives us permission to use and redistribute your contributions as part of the
project.

If you or your current employer have already signed the Google CLA (even if it
was for a different project), you probably don't need to do it again.

Visit <https://cla.developers.google.com/> to see your current agreements or to
sign a new one.

### Review our community guidelines

This project follows
[Google's Open Source Community Guidelines](https://opensource.google/conduct/).

## Contribution process

### Code reviews

All submissions, including submissions by project members, require review. We
use GitHub pull requests for this purpose. Consult
[GitHub Help](https://help.github.com/articles/about-pull-requests/) for more
information on using pull requests.

---

## Prerequisites

Before you begin local development, ensure you have the following installed on
your workstation:

- **Node.js**: v18.0.0 or higher (v20+ recommended)
- **npm**: v9.0.0 or higher
- **Python**: v3.10 or higher
- **Google Cloud SDK (`gcloud`)**: Latest version, authenticated with an account
  possessing Cloud Run, Discovery Engine, and Vertex AI permissions
- **Microsoft 365 Account**: With access to Word, PowerPoint, Excel, and Outlook
  (Desktop or Web)

---

## Local Development Setup

### 1. Clone the Repository

```bash
git clone https://github.com/GoogleCloudPlatform/gemini-for-microsoft-365.git
cd gemini-for-microsoft-365
```

### 2. Configure Environment Variables

Copy the template configuration file:

```bash
cp .env.example .env
```

Populate `.env` with your Google Cloud Project ID, region, service account, and
Microsoft Entra ID parameters.

### 3. Install Add-in Dependencies

```bash
cd microsoft-addin
npm install
cd ..
```

### 4. Start Local Development Server

To launch the Office Add-in local development server on `https://localhost:3000`:

```bash
cd microsoft-addin
npm run dev-server
```

### 5. Generate a Local Development Manifest

In a separate terminal, generate the development XML manifest:

```bash
python3 scripts/generate_manifest.py --dev --output microsoft-addin/manifest-dev.xml
```

You can now sideload `manifest-dev.xml` into Word, PowerPoint, Excel, or Outlook
(see [Deployment Guides](docs/deployment/)).

---

## Testing

### Running Frontend & Adapter Unit Tests

From the `microsoft-addin/` directory:

```bash
# Run the Excel and Taxonomy engine test suite (50+ tests)
npm test

# Run the Outlook email thread & calendar assist test suite (30+ tests)
npm run test:outlook
```

### Running Backend Auth Proxy Tests

From the `authproxy/` directory:

```bash
cd authproxy
pytest
cd ..
```

---

## Code Quality & Guardrails

### 1. Large JavaScript Files

Do **not** perform bulk rewrites or unverified regex replacements on
`slideBuilder.js`, `slideParser.js`, `chartRenderer.js`, or `taskpane.js`. When
editing these files, apply surgical, idempotent updates and run automated test
harnesses to prevent layout regressions.

### 2. Backward Compatibility

All Office host adapters (`WordAdapter`, `PPTAdapter`, `ExcelAdapter`,
`OutlookAdapter`) must gracefully handle scenarios where specific Office.js Rich
API features are unavailable (such as desktop webviews vs. Office on the Web)
using standard feature detection.

### 3. License Headers

Every new source file (`.js`, `.mjs`, `.css`, `.html`, `.py`, `.sh`, `.yaml`)
must include the standard Apache 2.0 license header (`Copyright 2026 Google LLC`).

---

## Pull Request Guidelines

1. **Branching**: Create a topic branch from `main` (e.g.,
   `feature/my-enhancement` or `fix/issue-description`).
2. **Commit Messages**: Use conventional commit formatting (e.g.,
   `feat(excel): add new chart formatting option`,
   `fix(outlook): preserve CC recipients on reply`).
3. **Verify Tests**: Ensure all automated test suites pass cleanly before
   opening a pull request:
   ```bash
   cd microsoft-addin && npm test && npm run test:outlook
   ```
4. **Documentation**: Update relevant specifications in `docs/specs/` or
   architecture decision records in `docs/adr/` if your change introduces new
   behaviors or configuration options.

---

## License

By contributing to this project, you agree that your contributions will be
licensed under the [Apache-2.0 License](LICENSE).
