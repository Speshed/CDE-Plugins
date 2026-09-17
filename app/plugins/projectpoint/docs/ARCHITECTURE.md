# Architecture

## Goal of the refactor

The source archive had business rules, HTTP requests, Excel parsing and Qt widgets in one `PP.py` file. This made every change risky because unrelated areas shared one namespace and one large `MainWindow` class.

The refactor separates the application into layers without intentionally changing the existing user workflow.

## Layers

### `projectpoint/config.py`
Connection defaults, predefined environments and `connection_profiles.json` access.

### `projectpoint/auth.py`
Authentication discovery and token acquisition for the supported ADFS and Keycloak/broker flows. The shared `requests.Session` also remains here.

### `projectpoint/api/`
Thin HTTP functions grouped by Project Point domain:

- `custom_fields.py`
- `project_stages.py`
- `content_types.py`
- `routes.py`
- `projects.py`
- `objects.py`
- `roles.py`

`common.py` contains shared request helpers and headers.

### `projectpoint/excel/`
Excel-specific parsing and workbook generation. This keeps pandas/openpyxl details out of the API layer.

### `projectpoint/ui/`
Qt presentation layer. `MainWindow` now handles only top-level orchestration. Large feature-specific method groups were moved into mixins under `ui/features/`.

### `projectpoint/resources/`
The large built-in permission dictionary was moved from Python source into JSON. It is data, not executable code, and is easier to update/diff in this form.

## API research notes

Raw HAR captures are not runtime dependencies and are not included in the clean working project. Confirmed findings needed for future changes are summarized in `docs/API_NOTES.md`.
