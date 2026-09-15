# Known issues found in the supplied source

These are pre-existing gaps in the uploaded `PP.py`, not functions lost during the refactor.

## Parameter import — write payload still unresolved

The Excel side is now implemented and tested: the generated parameter template is read back by `read_excel()`, its header row is auto-detected, columns are normalized, and `validate_row()` checks required fields, data types and list values.

The only remaining gap is `map_excel_row_to_payload()`. The exact `CustomFieldService/Create` write DTO — especially the structure used for values of type `Список` — is not present in the supplied HAR captures. The program therefore still raises `SourceGapError` before sending a guessed state-changing request.

To finish this safely, capture one successful parameter creation request from Project Point (preferably one normal text parameter and one `Список`) or provide an older source version containing the original mapper.

## Approval-route reference loading — mostly restored

The following missing helpers are now implemented:

- `get_all_projects` — mapped to the existing project lookup.
- `get_active_project_stages` — `ProjectStageService/GetAll`, filtered by `IsActive`.
- `get_object_structures_for_project` — existing object-structure query.
- `get_all_departments` — `Core/DepartmentsService/GetAll`, confirmed in the supplied Project Point frontend bundle; `Items` is unwrapped automatically.
- `get_all_disciplines` — the supplied frontend confirms the `Core/DisciplineService/` service root. The read-only `GetAll` action is used following the same reference-service contract and both `/ru/api/...` and `/api/...` variants are tried.

The disciplines `GetAll` action is a well-supported inference rather than a request directly observed in the supplied HAR. If a particular Project Point deployment exposes a different action, the route log will show the failed endpoint without modifying server data.

## Threaded Qt UI updates — fixed for import workers

Parameter, document-view, content-type, route, role and object-creation workers no longer re-enable their Qt buttons directly from background Python threads. Completion is marshalled back through Qt signals to the main window.
