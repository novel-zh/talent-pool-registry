# Setup, schema and deployment

## Source documents

Use native Google Sheets, rather than Excel files opened in the Google editor.
Import `examples/main.csv` into a `Main` tab and `examples/analysis.csv` into an `Analysis`
tab in two separate documents for a synthetic demo. All example identities are fictional.

`SCHEMA_` in `Code.gs` defines supported headings and aliases. The main table needs `Name`
and `Region`; the analysis table needs `ФИО участника` and `Регион`. Other mapped columns
provide optional fields. Each source must have exactly one matching header row among its first ten rows.

For all editable fields, create columns using the corresponding entries in `SCHEMA_.main.fields`
and `SCHEMA_.bkr.fields`. The minimal CSVs demonstrate connection and linking, not every optional field.

## Apps Script project

1. Add the two source files and enable the manifest in Project Settings.
2. Replace `appsscript.json` with the supplied manifest.
3. In Project Settings → Script Properties set:

| Property | Value |
| --- | --- |
| `MAIN_SPREADSHEET_URL` | Main Google Sheet URL or document ID |
| `BKR_SPREADSHEET_URL` | Analysis Google Sheet URL or document ID |
| `MAIN_TAB_NAME` | Main tab name; optional if discovery is unambiguous |
| `BKR_TAB_NAME` | Analysis tab name; optional if discovery is unambiguous |

The values in `config.example.json` are placeholders and are not loaded automatically.
Do not commit actual URLs, `.clasp.json`, account tokens or copied personnel data.

## Initialize

Run `initializeEmployeeIds` with editor access to both sources. It creates/fills the
`Сотрудник_ID` column and links unambiguous names. Existing IDs are preserved. Different people
must have distinct IDs; the same person in both sources must have the same ID.
Resolve ambiguous duplicate names manually, then rerun initialization.
Run `checkConnection` to verify loading.

## Deploy

Deploy → New deployment → Web app. Select **Execute as: User accessing the web app**.
Choose an audience permitted by your Google account/organization and restrict it to intended users.
Each user needs appropriate access to both underlying documents. Only editors can change the
corresponding source; creating a participant needs editor access to both sources.
The app is not a substitute for the Google documents' sharing controls.

After changing code, edit the deployment and select a new version. Refresh the browser.
Do not publish the working deployment URL in a public README containing personnel data.

## Operational notes

Start with synthetic documents or copies. Back up source documents before initializing IDs.
Revision checks reject stale edits; reload and reapply them after reviewing external changes.
Write recovery is best effort across two Sheets. If recovery reports incomplete changes,
inspect both sources and reconcile them before retrying.
