<div align="center">

# 🛠️ VBA-Actions
**Modular GitHub Actions to lint, test, document, and build VBA projects in CI/CD.**

[![GitHub Org](https://shields.io)](https://github.com)
[![License](https://shields.io)](https://github.com)

</div>

---

### 🚀 Why VBA-Actions?
Developing in VBA traditionally locks code inside binary Office files that are difficult to diff, review, and test automatically. **VBA-Actions** brings standard software engineering workflows to VBA—allowing you to maintain plain-text source files (`.bas`, `.cls`, `.frm`) in Git while automating style checks, unit testing, code coverage, API documentation, and `.xlam` add-in builds on GitHub.

### 🧰 Available Actions

| Action | Stage | Description |
| :--- | :--- | :--- |
| **[lint-vba](https://github.com/lint-vba)** | `Lint` | Scans VBA source files for whitespace, formatting, and capitalization issues so hand-written code matches the native MS Office VBA IDE style. |
| **[test-vba](https://github.com/test-vba)** | `Test` | Runs unit test functions against a collection of VBA source files, reports pass/fail results, and uploads code coverage metrics to Coveralls.io. |
| **[NaturalDocs-build-pages](https://github.com/NaturalDocs-build-pages)** | `Docs` | Generates HTML API documentation from source comments using NaturalDocs and deploys it directly to GitHub Pages. |
| **[build-excel-addin](https://github.com/build-excel-addin)** | `Build` | Bundles a directory of VBA source code files into a compiled `.xlam` Excel add-in for automated releases. |

---

### ⚡ End-to-End CI/CD Workflow Example

You can chain all four actions together in `.github/workflows/ci-release.yml` to validate pull requests, update your documentation site on merges to `main`, and automatically attach a compiled `.xlam` add-in on version tags:

```yaml
name: VBA CI/CD Pipeline

on:
  push:
    branches: [ main ]
    tags: [ 'v*' ]
  pull_request:
    branches: [ main ]

jobs:
  lint-and-test:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout source
        uses: actions/checkout@v4

      - name: Lint VBA source files
        uses: Vba-actions/lint-vba@v1

      - name: Run VBA unit tests & upload coverage
        uses: Vba-actions/test-vba@v1
        env:
          GITHUB_TOKEN: \${{ secrets.GITHUB_TOKEN }}

  build-docs:
    needs: lint-and-test
    if: github.ref == 'refs/heads/main'
    runs-on: windows-latest
    permissions:
      contents: write
      pages: write
      id-token: write
    steps:
      - name: Checkout source
        uses: actions/checkout@v4

      - name: Generate & deploy NaturalDocs to GitHub Pages
        uses: Vba-actions/NaturalDocs-build-pages@v1

  build-and-release:
    needs: lint-and-test
    if: startsWith(github.ref, 'refs/tags/v')
    runs-on: windows-latest
    permissions:
      contents: write
    steps:
      - name: Checkout source
        uses: actions/checkout@v4

      - name: Bundle source into .xlam add-in
        uses: Vba-actions/build-excel-addin@v1

      - name: Upload .xlam to GitHub Release
        run: gh release upload "\${{ github.ref_name }}" ./dist/*.xlam --clobber
        env:
          GITHUB_TOKEN: \${{ secrets.GITHUB_TOKEN }}
```
*(Note: Adjust the `with:` inputs, branch/tag conditions, and runner OS—e.g., `ubuntu-latest` if any of your actions are cross-platform—to match your `action.yml` configurations.)*

---

### 🗺️ Roadmap & Wishlist
* **Static Analysis (`analyze-vba`)** *(Planned)* — Catch undeclared variables (`Option Explicit` enforcement), unused procedures/variables, unreachable code, and type-coercion pitfalls before runtime.
* **Dependency & Reference Manager (`restore-vba-refs`)** *(Planned)* — Resolve COM/TypeLib GUIDs (`Scripting`, `ADODB`, `MSXML2`) and fetch pinned third-party `.bas`/`.cls` libraries from a `vba-project.toml` manifest before testing or building.

---

### 🤝 Contributing & Feedback
Have an idea for a new VBA CI/CD action or found a bug in an existing workflow? Open an issue or pull request in the relevant repository!
