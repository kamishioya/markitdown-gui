# MarkItDown GUI

MarkItDown GUI is a Windows-oriented PySide6 desktop shell for converting local PDF, DOCX, XLSX, and HTML files to Markdown.

For XLSX files, the GUI exports the workbook to PDF through Microsoft Excel on Windows and then runs the existing PDF-to-Markdown path. If GitHub Copilot CLI post-processing is enabled, that shaping step still runs last.

This repository contains the GUI layer only. The conversion engine comes from the published `markitdown` package maintained by Microsoft.

This package is intentionally narrow in scope:

- Local files only
- Supported formats: PDF, DOCX, XLSX, HTML, HTM
- Plugins disabled
- No OCR plugin wiring
- No Azure Document Intelligence wiring

## Features

- Multi-file selection and drag-and-drop
- Batch conversion to Markdown
- Optional GitHub Copilot CLI post-processing that reshapes the generated Markdown in place
- XLSX conversion via internal Excel-to-PDF export before Markdown generation
- Output folder selection
- Japanese UI by default with an English language switch
- Optional overwrite and keep-data-URI behavior
- Conversion log and Markdown preview
- PyInstaller spec for Windows distribution

## Development Setup

From the repository root:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[build]"
```

## Run The GUI

From the repository root after installation:

```bash
markitdown-gui
```

Or directly:

```bash
python -m markitdown_gui
```

The UI starts in Japanese by default. You can switch the display language to English from the language selector in the main window, and the selection is persisted for the next launch.

If you enable GitHub Copilot CLI post-processing, the GUI calls `copilot -p` after the base conversion and overwrites the generated Markdown with the shaped result. Install GitHub Copilot CLI separately, complete the CLI login flow once, and then either leave the command field empty for auto-detection or point the GUI at a specific `copilot` executable. The desktop app now also includes an in-app setup dialog under the Settings menu so you can auto-detect the command, open the installation guide, and launch a CLI window for `/login`.

The GUI intentionally ignores the VS Code-internal wrapper under AppData/Roaming/Code/User/globalStorage/github.copilot-chat/copilotCli unless a real GitHub Copilot CLI installation is available behind it. This avoids getting stuck in the wrapper's interactive install flow.

For `.xlsx` files, the GUI launches Microsoft Excel through Windows PowerShell, exports the workbook to a temporary PDF, and then converts that PDF through MarkItDown. This means `.xlsx` conversion now depends on a local Excel installation that can be automated through COM.

## Build A Windows Distribution

Run the build wrapper from the repository root:

```bash
build-release.cmd
```

The bundled application is created under `release/MarkItDownGUI`.
If the repository contains a top-level `LICENSE` file, the distribution includes it alongside the executable.

Launch only this executable:

```bash
release/MarkItDownGUI/MarkItDownGUI.exe
```

Do not launch anything under `build/`. PyInstaller places temporary intermediate executables there, and they are not distribution-ready.

## GitHub Release Workflow

This repository includes a GitHub Actions workflow that builds a Windows release when you push a tag such as `v0.1.0`.

- `push` on tags matching `v*`: runs tests, builds the PyInstaller package, creates a zip archive, and publishes a GitHub Release.
- `workflow_dispatch`: runs the same build and uploads the zip as a workflow artifact without publishing a release.

## Upstream MarkItDown

- Upstream project: https://github.com/microsoft/markitdown
- This GUI depends on `markitdown[pdf,docx,xlsx]` at runtime and is not an official Microsoft project.
- If this repository only depends on the published package and does not copy upstream source code, keeping an upstream link in the README is recommended but not strictly required by the MIT license.
- If you copied or modified source files from the Microsoft repository, keep the original MIT license notice for the copied portions and document that relationship in this repository.

## Notes

- This package depends on `markitdown[pdf,docx,xlsx]`.
- `.xlsx` conversion depends on Microsoft Excel being installed on the target Windows machine.
- GitHub Copilot CLI is not bundled into the desktop application. Install it separately with `winget install GitHub.Copilot` or `npm install -g @github/copilot` before enabling the post-processing option.
- The PyInstaller spec explicitly collects `magika` assets because MarkItDown uses them at runtime for file identification.
- If you want OCR, plugins, or cloud-backed conversion flows, add them after the base desktop flow is stable.
