# Project Point API notes

Useful API findings extracted from the supplied HAR captures and bundled frontend sources. Raw HAR files are intentionally not part of the working project because they are large and may contain session/authentication metadata.

## Document views / project stages

Confirmed requests from `Создание видов документов.har`:

- `GET /ru/api/Core/ProjectStageService/GetAll`
- `POST /ru/api/Core/ProjectStageService/Create`
- `POST /ru/api/Core/ProjectStageService/Update`

Observed create payload:

```json
{
  "Title": "Тестовый",
  "Code": "001",
  "DisplayParameters": 0,
  "IsActive": true
}
```

Observed update payload adds the existing `Id`:

```json
{
  "Id": "<existing id>",
  "Title": "Тестовый",
  "Code": "001",
  "DisplayParameters": 1,
  "IsActive": true
}
```

Implemented in:

- `projectpoint/api/project_stages.py` — GetAll/Create/Update.
- `projectpoint/excel/project_stages.py` — Excel parsing and validation.
- `projectpoint/ui/features/project_stages.py` — separate **Виды документов** tab.

Import behavior: match by `Code` (case-insensitive for duplicate protection), create missing records, update changed records, skip identical records.

## Route reference services

The supplied Project Point frontend bundle confirms:

- service root `Core/DisciplineService/`;
- `Core/DepartmentsService/GetAll` for organizational units, with records under `Items`.

For disciplines the application uses the read-only `DisciplineService/GetAll` action as a service-contract inference and tries locale/non-locale URL variants. This action was not directly captured in the supplied HAR.
