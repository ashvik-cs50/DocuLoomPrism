# DocuLoom Prism

`DocuLoom Prism` is a desktop AI-style document insight app that turns files into clear, presentation-ready charts. It accepts common business documents, extracts usable numeric data, supports linked image references, and exports polished PowerPoint decks that are ready to share.

[![Download Release](https://img.shields.io/badge/Download-Latest%20Release-1f6f63?style=for-the-badge)](https://github.com/ashvik-cs50/DocuLoomPrism/releases)

## Download

Want the compiled app instead of running from source?

[Download the latest release from GitHub](https://github.com/ashvik-cs50/DocuLoomPrism/releases)

## Why It Feels Product-Ready

- Unique app branding with a stronger public-facing name
- Desktop UI designed for demos, publishing, and GitHub presentation
- Support for `PDF`, `DOCX`, `XLSX`, `XLS`, `XLSM`, `PPTX`, `CSV`, and `TXT`
- Optional URL fetching for direct images and supported document files
- Graph generation with topic labels, numbering, and readable titles
- PowerPoint export with chart slides, summary slides, and linked-image slides
- Automatic Windows executable build pipeline for GitHub Actions
- Logo-ready setup for your `prizm.ico` file and future `logo.png`

## Product Flow

1. Open a document.
2. Let the app detect chart-ready data.
3. Choose a dataset and numeric value column.
4. Select a graph type from the dropdown.
5. Optionally fetch an image or supported document from a direct URL.
6. Generate a polished graph.
7. Export a presentation-ready PowerPoint.

## Supported Charts

- Bar Chart
- Line Chart
- Horizontal Bar Chart
- Area Chart
- Scatter Plot
- Pie Chart
- Histogram

## Repository Structure

- `doculoom_prism.py`: main desktop application
- `requirements.txt`: runtime dependencies
- `build_requirements.txt`: packaging dependency for compiled builds
- `build_release.ps1`: local Windows build script for `.exe` output
- `BUILDING.md`: packaging and GitHub release instructions
- `.github/workflows/build-windows.yml`: GitHub Actions build pipeline
- `assets/README.md`: where to place your future logo files

## Local Setup

```bash
pip install -r requirements.txt
```

## Run The App

```bash
python doculoom_prism.py
```

If `python` is not available in your terminal on Windows, use:

```bash
py doculoom_prism.py
```

## Build A Downloadable Windows App

Install build dependencies:

```bash
pip install -r build_requirements.txt
```

Build the executable:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_release.ps1
```

Compiled output:

```text
dist/DocuLoomPrism.exe
```

## Add Your Logo Later

When you are ready, place these files in `assets/`:

- `prizm.ico`: used for the Windows `.exe`
- `logo.png`: used for the app window icon when supported

The app and build process already detect these automatically.

## GitHub Download Story

This repo now includes a GitHub Actions workflow that builds the Windows app automatically on `main` or `master`, and also supports manual workflow runs.

That means people can:

- open the repository on GitHub
- trigger or inspect the workflow
- download the compiled Windows artifact
- later download a release build once you upload the `.exe` to GitHub Releases

## Notes

- `PDF`, `DOCX`, and `PPTX` work best when the file contains tables or obvious label-value pairs.
- URL fetching works best with direct file links, not general web pages.
- `Pie Chart` requires values greater than zero.
- Larger datasets are trimmed for some chart types so labels stay clean and readable.

## Ready For Next Steps

This project is now set up for:

- branding with your future logo
- local `.exe` builds
- GitHub-hosted build artifacts
- a more publishable product presentation
