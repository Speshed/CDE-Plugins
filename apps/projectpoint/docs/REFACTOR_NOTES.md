# Refactor notes

## Changed

- Split the original ~7700-line `PP.py` into a Python package.
- Split `MainWindow` into feature mixins.
- Split HTTP API functions by domain.
- Split Excel parsers/template builders from HTTP/UI code.
- Moved the large built-in role permission dataset to JSON.
- Removed the transitional root `PP.py`; `main.py` is now the only application entry point.
- Removed `__pycache__` and `.pyc` artifacts from the deliverable.
- Moved sample workbooks to `examples/`.
- Removed raw HAR captures from the clean working tree after documenting confirmed API findings.
- Decoded escaped `#Uxxxx` filenames to readable Unicode names.
- Added a single `requirements.txt`, `.gitignore`, `run.bat`, and an updated `build.bat`.
- Updated PyInstaller build to include required JSON resources.

## Intentionally not changed

- Existing tab behavior and worker logic.
- Existing API payloads and endpoints.
- Existing authentication workflows.
- Existing Excel import schemas.
- The proposed "document views" button is not wired into the UI yet; that should be the next isolated feature change.
