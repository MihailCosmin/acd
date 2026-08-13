# acd
ALTHOM Codebase

# INSTALL
pip install git+https://github.com/MihailCosmin/acd


### INSTALL WITH VERSION
python -m pip uninstall -y acd

pip wheel git+https://github.com/MihailCosmin/acd -w dist

##### Install with all dependencies
python -m pip install --force-reinstall --no-index dist/acd-0.0.3.1-py3-none-any.whl

##### Install without dependencies - only update acd
python -m pip install -U --no-deps dist/acd-0.0.3.1-py3-none-any.whl

pip freeze | findstr acd   # should show: acd==0.0.3.1

# LINUX (Ubuntu 22.04) NOTES
On Windows, this package uses the bundled binaries under `acd/3rd` (poppler,
tesseract, ccache) and Word/COM automation for `.docx` -> `.pdf` conversion.
On Linux those are not available, so the code instead relies on system
packages being installed and available on `PATH`:

    sudo apt install poppler-utils tesseract-ocr libreoffice

- `poppler-utils` replaces the bundled poppler binaries (PDF -> image conversion).
- `tesseract-ocr` replaces the bundled tesseract binary (OCR fallback engine).
- `libreoffice` (`soffice`) replaces Word/COM automation for `.docx` -> `.pdf` conversion.

CGM -> SVG conversion (`cgm2svg`, `cgm2clearcgm`) relies on Windows-only
executables bundled in `acd/3rd` with no Linux equivalent available; those
features remain Windows-only.