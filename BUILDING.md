# Building DocuLoom Prism

This project includes a Windows build script so you can produce a downloadable desktop app for GitHub releases.

## Local Build

1. Install Python and dependencies:

```bash
pip install -r requirements.txt
pip install -r build_requirements.txt
```

2. Put your icon file at:

```text
assets/prizm.ico
```

3. Run the Windows build script:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_release.ps1
```

4. Your compiled app will be created at:

```text
dist/DocuLoomPrism.exe
```

If the old `.exe` would not launch, rebuild it after the latest packaging changes. The build script now bundles the dynamically imported document libraries explicitly for PyInstaller and uses the working `python` interpreter instead of the broken `py` launcher.

## Debug Build

If the normal GUI `.exe` still does not open, build the debug version so a console window shows the actual startup error:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_debug.ps1
```

Debug output:

```text
dist/DocuLoomPrism-Debug.exe
```

## GitHub Download Flow

- Push the repo to GitHub.
- The included GitHub Actions workflow builds the Windows app automatically.
- Download the compiled artifact from the workflow run.
- Once you start tagging releases, you can also attach `DocuLoomPrism.exe` to GitHub Releases.

## Logo Setup

- `assets/prizm.ico` is used for the compiled `.exe` icon when present.
- `assets/logo.png` is used by the app window when present.
- If no logo is added yet, the app still builds successfully.
