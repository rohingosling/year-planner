# 📅 Year Planner Generator

![Python](https://img.shields.io/badge/Python-3.12%2B-3776AB?style=flat&logo=python&logoColor=white)
![python-docx](https://img.shields.io/badge/python--docx-1.2-3776AB?style=flat)
![PDF via Word](https://img.shields.io/badge/PDF%20via-Microsoft%20Word-2B579A?style=flat&logo=microsoftword&logoColor=white)
![PyInstaller](https://img.shields.io/badge/Packaging-PyInstaller-FFCA28?style=flat)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-44CC11?style=flat)
![License](https://img.shields.io/badge/License-MIT-green?style=flat)
![Version](https://img.shields.io/badge/version-1.2-orange?style=flat)

Generate a printable **Year Planner** as a Microsoft Word document (or PDF on Windows and macOS), optimized for duplex printing.

Run it two ways:

- **Download the standalone Windows executable** from [Releases](https://github.com/rohingosling/year-planner/releases/latest) — no Python required.
- **Clone and run** the Python source on **Windows**, **macOS**, or **Linux**.

A `.docx` is always produced. A `.pdf` is an optional output using the `--pdf` argument, that requires Microsoft Word — supported on **Windows** and **macOS**.

> **Note:** The **macOS** path is wired up but untested; see note below. **Linux** is DOCX-only.

## ✨ What it generates

A complete, print-ready planner with a minimal black/white/grayscale aesthetic and a strict recto/verso layout (mirror margins and a binding gutter):

- Cover page with lost-and-found contact info
- Instructions page
- Year-at-a-glance calendars (current year + next year)
- Table of Contents with pre-calculated page numbers
- Goals page
- Backlog section for unscheduled tasks
- Week Planner with ISO 8601 week numbering
- Monthly sections with daily spreads (×12)
- Terms and Definitions
- Graph-paper pages
- Rear cover (blank, reserved for future barcode/branding)

## 🚀 Quick Start

Download [`year-planner.exe`](https://github.com/rohingosling/year-planner/releases/latest) and run it
from the Windows Command Prompt or PowerShell:

**Examples:**
```cmd
REM Current year, DOCX into the current directory
year-planner.exe

REM A specific year
year-planner.exe --year 2027

REM Choose the output name
year-planner.exe --year 2027 --filename my-planner

REM Also produce a PDF (Windows or macOS + Microsoft Word only)
year-planner.exe --year 2027 --pdf
```

The planner's style and layout are fixed in the build. The only thing you choose is the **year**.

## 📑 CLI reference

| Argument | Default | Description |
|----------|---------|-------------|
| `--year YEAR` | current year | Planner year (validated 1000–9999) |
| `--filename NAME` | `year-planner-<year>` | Output filename or path (in the current directory). A `.docx`/`.pdf` extension is optional; the correct one is applied per format |
| `--pdf` | off | Also produce a PDF. Requires Microsoft Word (**Windows**, or **macOS** — untested). Fails fast if Word isn't detected |
| `--version` | | Print the version and exit |
| `--help` | | Show help and exit |

### 📄 PDF / Word note

PDF conversion uses Microsoft Word via `docx2pdf`:

- **Windows** with Word installed — supported and tested.
- **macOS** with Microsoft Word for Mac installed — supported via docx2pdf's AppleScript path, but **not tested
  by the author**. It should work; please report any issues.
- **Linux** — not supported (docx2pdf has no Linux path and Word does not run natively there); DOCX only.

On any platform without Word, open the generated DOCX in Word, LibreOffice, or Google Docs and export to PDF.

## 🐍 Run from source

Requires **Python 3.12+**.

```bash
# Clone
git clone https://github.com/rohingosling/year-planner.git
cd year-planner

# Create and activate a virtual environment
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

# Install runtime dependencies
pip install -r requirements.txt

# Generate (current year)
python src/main.py --year 2027
```

On Windows you can also use the launcher:

```cmd
run.bat --year 2027
```

## 🛠️ Build the executable

Building the standalone Windows `.exe` needs the dev/build dependencies:

```cmd
.venv\Scripts\python.exe -m pip install -r venv_requirements.txt
.venv\Scripts\python.exe scripts\build.py
```

`build.py` stamps the version and build date, compiles via PyInstaller using `year-planner.spec`, and runs a
DOCX-only smoke test. The result is `dist/year-planner.exe`.

Equivalent raw PyInstaller invocation (run from the repo root):

```cmd
pyinstaller --onefile --console --name year-planner ^
  --add-data "config/config.yaml;config" ^
  --add-data "assets/images/instructions.png;assets/images" ^
  --add-data "assets/images/graph_paper_37x56_15_50_2161x3295px.png;assets/images" ^
  src/main.py
```

(The `--add-data` separator is `;` on Windows and `:` on macOS/Linux.)

## ⚙️ Configuration model

`config/config.yaml` is the **single source of truth** for the document's layout and style, but it is **build-time configuration**, not a user setting: it is bundled into the executable at build time and read live when running from source. End users do not edit it, the only user-facing configuration is `--year`. To change the fixed style, edit `config.yaml` and rebuild. (A hidden `--config FILE` flag exists for power users running from source.)

## 🖨️ Printing

For best results:

1. Print duplex (two-sided), flip on the long edge.
2. Use A4 paper.
3. Set print scaling to 100% (no fit-to-page).
4. Bind on the left edge.

## 📁 Project structure

```
year-planner/
├── src/                  # Application source
│   ├── main.py           # CLI entry point
│   ├── config.py         # YAML configuration loader
│   ├── document.py       # Document init, page/section breaks
│   ├── paths.py          # Resource/cache/output path resolution
│   ├── sections/         # One module per document section
│   └── utils/            # Shared helpers (styles, tables, grid images)
├── config/config.yaml    # Build-time configuration (bundled into the exe)
├── assets/images/        # Bundled images (instructions, graph-paper grid)
├── scripts/build.py      # PyInstaller build + smoke test
├── year-planner.spec     # PyInstaller build spec
├── requirements.txt      # Runtime dependencies
└── venv_requirements.txt # Dev/build dependencies
```

## 📜 License

Released under the [MIT License](LICENSE). Third-party components and their licenses are listed in
[THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md).

## ✍️ Author

[Rohin Gosling](https://github.com/rohingosling)
