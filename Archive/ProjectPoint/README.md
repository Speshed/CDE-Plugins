# ProjectPoint

Desktop utility for Project Point automation and Excel-based imports.

## Run

Install dependencies once:

```bat
pip install -r requirements.txt
```

Then either run:

```bat
run.bat
```

or:

```bat
python main.py
```

`main.py` is the **only application entry point**.

## Build EXE

```bat
build.bat
```

The executable is created as `dist\ProjectPoint.exe`.

## Project structure

```text
ProjectPoint/
├─ main.py                    # only entry point
├─ connection_profiles.json  # connection profiles
├─ requirements.txt          # runtime + build dependencies
├─ run.bat
├─ build.bat
├─ run_tests.bat
│
├─ projectpoint/
│  ├─ auth.py                 # ADFS / Keycloak authentication
│  ├─ config.py               # configuration and profiles
│  ├─ styles.py               # Qt stylesheet
│  ├─ api/                    # Project Point API by domain
│  ├─ excel/                  # Excel parsing/templates
│  ├─ ui/                     # Qt UI
│  │  └─ features/            # one module per functional section
│  └─ resources/              # bundled data files
│
├─ examples/                  # sample Excel files
├─ tests/                     # smoke tests
└─ docs/                      # architecture/API notes/known issues
```

There is intentionally no root `PP.py`, no duplicate copy of the old monolith, and no second application launcher.

## Functional sections

- `parameters.py` — parameters/custom fields (source API gap remains; see `docs/KNOWN_ISSUES.md`).
- `project_stages.py` — document views: Excel template, create/update/skip.
- `content_types.py` — content/document types.
- `routes.py` — approval routes.
- `roles.py` — roles and permissions.
- `objects.py` — construction-object hierarchy.
- `connection.py` — connection/authentication.

## Current status

- Structural refactor is complete.
- **Document views / project stages are implemented** as a separate tab using the confirmed GetAll/Create/Update API.
- Generated document-type templates are now directly importable: the importer auto-detects the real table-header row below the template title.
- Route reference loading now includes organizational units and a conservative read-only disciplines lookup.
- The only known source-level blocker is the original parameter-import payload gap documented in `docs/KNOWN_ISSUES.md`.
