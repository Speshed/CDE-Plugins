# Source audit

## Uploaded project

The uploaded archive contained one main Python module, `PP.py`, with **7,727 lines**, plus a tiny `main.py`, build script, connection profile JSON, three Excel workbooks, two HAR captures and a generated `__pycache__` file.

`PP.py` contained:

- 79 top-level functions;
- 4 Qt classes (`LogSignal`, `ObjectsLoadSignal`, `LogDialog`, `ConnectionDialog`) plus `MainWindow`;
- 85 methods in `MainWindow`;
- a large embedded role-permission dataset;
- authentication, HTTP API, Excel and UI logic in the same namespace.

## Result after refactor

All 79 original top-level function names are retained in the new package and all 85 original `MainWindow` methods are retained. The large role-permission dataset is stored as JSON instead of Python source.

The largest remaining feature module is the route workflow. It is intentionally kept as one feature unit in this pass because much of its parsing/matching logic is nested inside `_run_routes_worker`; splitting that logic further would be a behavior-changing refactor and should be done with route-specific regression fixtures.

## Validation performed

- `python -m compileall` succeeds for the full refactored source tree.
- Core modules import successfully in the working Python environment.
- Static bytecode checks found no unresolved global references introduced by module splitting.
- Role-permission JSON loads 196 source records.
- Role template generation and parser round-trip succeeds (130 visible permission rows in the default template).
- Object import workbook parsing succeeds on the `Импорт` sheet from the supplied example.
- Parameter/content-type/route template workbook generation was smoke-tested where the original source provides complete implementations.
- The environment used for this refactor does not have PySide6 installed, so a real Qt window launch was not executed; UI modules were import-checked using lightweight Qt stubs in addition to syntax compilation.

## Original-source gaps

See `KNOWN_ISSUES.md`. Those missing helpers existed as calls in the uploaded source but had no definitions in either `PP.py` or its matching compiled `.pyc`.
