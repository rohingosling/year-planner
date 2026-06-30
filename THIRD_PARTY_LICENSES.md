# Third-Party Licenses

The Year Planner Generator is distributed under the [MIT License](LICENSE). It builds on the third-party
components listed below. This notice is provided for transparency and to satisfy the attribution requirements of
those components.

## Runtime dependencies

These packages are required to run the generator from source and are bundled into the standalone Windows
executable (`dist/year-planner.exe`). All use permissive licenses compatible with redistribution.

| Package | Version | License | Project |
|---------|---------|---------|---------|
| python-docx | 1.2.0 | MIT | https://github.com/python-openxml/python-docx |
| PyYAML | 6.0.3 | MIT | https://pyyaml.org/ |
| Pillow | 12.2.0 | MIT-CMU (HPND) | https://github.com/python-pillow/Pillow |
| lxml | 6.0.2 | BSD-3-Clause | https://lxml.de/ |
| docx2pdf | 0.1.8 | MIT | https://github.com/AlJohri/docx2pdf |
| typing_extensions | 4.15.0 | PSF (Python Software Foundation License) | https://github.com/python/typing_extensions |

`docx2pdf` is only exercised by the optional `--pdf` add-on. On Windows it uses `pywin32`
(PSF-2.0 / OSI-approved) to drive Microsoft Word via COM; `pywin32` is pulled in transitively and is not a
direct dependency of the generator.

## Packaging

| Component | License | Notes |
|-----------|---------|-------|
| PyInstaller | GPL-2.0-or-later **with bootloader exception** | The bootloader exception explicitly permits distributing the frozen application (this executable) under any license of your choice, including the MIT License used here. PyInstaller itself is a build-time tool and no PyInstaller source is included in the shipped binary beyond the exception-covered bootloader. |

## External programs

| Component | License | Notes |
|-----------|---------|-------|
| Microsoft Word | Proprietary (user-supplied) | Not distributed with this project. Used only by the optional `--pdf` add-on on Windows, via the user's own installation, to convert the generated DOCX to PDF. |

## Development-only (not shipped)

| Package | Version | License | Notes |
|---------|---------|---------|-------|
| PyMuPDF | ≥ 1.24.0 | AGPL-3.0 | Used **only** by the maintainer utility `scripts/pdf_to_png.py`. It is listed in `venv_requirements.txt`, is **not** in `requirements.txt`, is **not** imported by the generator, and is **not** bundled into the executable. The AGPL therefore does not apply to the distributed program. |
| PyInstaller, pywin32 | (latest) | see above | Build/runtime tooling listed in `venv_requirements.txt`. |

## Full license texts

The full text of each license is available from the corresponding project page linked above, and is included in
each package's distribution within your Python environment (typically under `*.dist-info/`).
