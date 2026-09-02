# Official Google Sheets API Integration

Use the official Google Sheets API for real Google Sheets work. Do not use a
third-party Sheets MCP server or treat an uploaded XLSX file as the primary
Google Sheets integration.

The implementation provides user-local OAuth, a structural read-only inspector,
and reusable read/write tooling. Every write must follow the safety workflow
below and be read back before it is considered successful.

## Required workflow

**READ**
→ **INVENTORY**
→ **PLAN**
→ **WRITE**
→ **READ BACK**
→ **VERIFY**

1. **READ** — parse the supplied URL or ID and retrieve spreadsheet metadata.
2. **INVENTORY** — confirm exact sheet titles and IDs; inspect relevant values,
   formulas, validations, checkboxes, formatting, protections, filters,
   conditional formats, named ranges, merges, charts, and hidden sheets.
3. **PLAN** — define the smallest exact ranges and properties that must change.
4. **WRITE** — use targeted value requests or `spreadsheets.batchUpdate` with
   explicit field masks. Preserve everything outside the planned scope.
5. **READ BACK** — retrieve every changed range and affected property again.
6. **VERIFY** — compare actual values, formulas, controls, formatting, and
   dependent outputs with the plan. A successful write response is not proof.

Never claim a Google Sheets change succeeded until the changed content was read
back and verified.

## OAuth 2.0 desktop authentication

Use OAuth 2.0 Desktop App authorization for the user's own Google account.
Service accounts are not the default user workflow.

User-local storage:

```text
%LOCALAPPDATA%\CursorExcelSkill\google\
├── oauth-client.json
├── token-sheets-readonly.json
└── token-sheets-readwrite.json   # future write integration only
```

- `oauth-client.json` is the OAuth Desktop client configuration downloaded
  manually from Google Cloud.
- Token files contain access and refresh token state and granted scopes.
- Client configuration and tokens remain separate.
- These files must never be stored in the repository, source code, `SKILL.md`,
  prompts, logs, or terminal output.
- Do not create, guess, or download OAuth credentials automatically.
- Never print client secrets, access tokens, or refresh tokens.

`scripts/google_auth.py`:

- fails clearly when the user-local OAuth client configuration is missing;
- supports explicit first-time browser authorization;
- refreshes an expired token on later runs when a refresh token is available;
- writes tokens atomically to the user-local directory;
- uses separate token files for read-only and future read/write grants.

Show the expected local paths without authenticating:

```bash
python scripts/google_auth.py --show-paths
```

After the manual Google Cloud setup, authorize read-only access:

```bash
python scripts/google_auth.py --authorize-readonly
```

Browser authorization must be explicit. Inspection does not launch a browser
unless `--authorize` is supplied.

Authorize separate read/write access only when a write task is explicitly
requested:

```bash
python scripts/google_auth.py --authorize-readwrite
```

This writes only `token-sheets-readwrite.json`; it does not overwrite the
read-only token.

## OAuth scopes

Current read-only inspector:

```text
https://www.googleapis.com/auth/spreadsheets.readonly
```

Future create/edit tooling:

```text
https://www.googleapis.com/auth/spreadsheets
```

Request only the scope required by the active tool. Do not request Google Drive
scope merely to inspect or edit a spreadsheet when its URL or ID is already
known. A future write tool must use the separate read/write token and may
require explicit reauthorization.

## Read-only inspection

Run:

```bash
python scripts/google_sheets_inspect.py \
  "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit"
```

The tool accepts a Google Sheets URL or spreadsheet ID, extracts and validates
the ID, obtains read-only OAuth credentials, and uses only:

- `spreadsheets.get`;
- `spreadsheets.values.get`.

It reports spreadsheet and sheet metadata, dimensions, frozen rows and columns,
values-based used ranges, sampled values, formulas, validations, checkboxes,
protected ranges, merges, filters and filter views, conditional formats, named
ranges, charts, slicers, hidden sheets, and relevant sampled cell formatting.

For large used ranges, detailed grid inspection is capped and the sampled range
is reported. Values-based used-range detection does not include cells that
contain only formatting. The Sheets API exposes embedded charts and slicers but
does not comprehensively expose every embedded object type.

Allow first-time browser authorization only when explicitly intended:

```bash
python scripts/google_sheets_inspect.py SPREADSHEET_ID --authorize
```

The inspector contains no write, clear, update, or batch-update operation.

## Read/write tooling

`scripts/google_sheets_write.py` provides reusable operations for:

- creating spreadsheets and creating or deleting sheets;
- reading, writing, batch-writing, and appending values or formulas;
- formatting cells with explicit field masks;
- resizing rows and columns and freezing panes;
- dropdown validation and native checkbox validation;
- conditional formatting and warning-only protected ranges;
- charts and general `spreadsheets.batchUpdate` requests;
- immediate value readback after writes and appends.

For every write task:

- use `spreadsheets.values.update` or `spreadsheets.values.batchUpdate` for
  targeted value/formula changes;
- use `spreadsheets.batchUpdate` for structural, validation, formatting,
  protection, filter, named-range, and chart operations;
- use explicit field masks on requests such as `UpdateCellsRequest`,
  `RepeatCellRequest`, and property updates so unrelated fields are not reset;
- separate read-only inspection from write execution;
- record the exact planned spreadsheet ID, sheet ID, ranges, request types, and
  field masks before writing;
- re-read changed ranges and affected metadata after every write;
- reconcile dependent summaries and charts with representative inputs.

Never clear or replace a whole sheet when a smaller operation is sufficient.

## Lessons confirmed by the first live write

- The Sheets scope alone can create and edit a spreadsheet; Google Drive scope
  is not required when file discovery or Drive operations are not needed.
- Keep read-only and read/write grants in separate token files.
- Spreadsheet locale affects formula parsing. For a `ru_RU` spreadsheet,
  semicolon-separated formulas were accepted and must still be read back using
  both `FORMULA` and `FORMATTED_VALUE` render modes.
- A `spreadsheets.batchUpdate` request is atomic. If one request is invalid,
  correct the request and resume the same spreadsheet instead of creating a
  duplicate.
- Chart properties must be placed in the correct chart-type schema. For
  example, legend placement belongs in `basicChart.legendPosition` or
  `pieChart.legendPosition`, not at the top of `EmbeddedChartSpec`.
- Native checkbox validation can materialize `FALSE` values and extend the
  values-based used range. Apply checkboxes to a practical initial entry area,
  and ensure formulas exclude blank ledger rows using a real key such as Date.
- Open-ended source ranges make dropdown lists expandable. Validation targets
  should still use the smallest practical range.
- Use warning-only protected ranges for calculated columns, analytical sheets,
  and dashboards when accidental edits should be discouraged without changing
  ownership or editor permissions.
- Verify formulas, validations, checkbox condition types, conditional-format
  rules, chart objects, frozen rows, filters, and protected ranges from a fresh
  API read after the write.
- An empty production template can validate chart objects and source ranges,
  but not real-data chart density or readability. Perform a later populated
  regression test before calling data-driven dashboard behavior fully
  validated.

## Permanent safety rules

- Never guess spreadsheet IDs or sheet names.
- Accept Google Sheets URLs and reliably extract the spreadsheet ID.
- Inspect an existing spreadsheet before modification.
- Prefer the smallest possible write or update scope.
- Preserve unrelated formulas, formatting, validations, checkboxes, charts,
  protections, filters, filter views, merges, and named ranges.
- Use field masks with `batchUpdate` requests so unrelated properties are not
  overwritten.
- Use `batchUpdate` for structural and formatting operations where appropriate.
- Re-read changed ranges after every write operation.
- Never claim success until the write was read back and verified.
- Never put OAuth secrets or tokens in repository files, source code,
  `SKILL.md`, logs, prompts, or terminal output.

## Google Sheets is not XLSX

Design Google Sheets natively rather than blindly translating Excel behavior:

- use native Google Sheets checkboxes where appropriate;
- use native Sheets data-validation rules;
- verify formulas against Google Sheets syntax and locale behavior;
- design protections for Sheets permissions and protected ranges;
- use the Sheets chart API and its object model;
- do not assume Excel Tables, structured references, calculated columns,
  worksheet protection, chart XML, or Excel rendering behavior exists in
  Google Sheets;
- do not use an XLSX upload to Drive as a substitute for a native Sheets design.

## Official Python dependencies

Install only when real authorization or API access is requested:

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
```

These are the official Google Python client and authentication libraries. No
Google package is required merely to read this architecture or run CLI help.

## Manual Google Cloud setup

1. Sign in to Google Cloud Console with the account that will own the OAuth
   project.
2. Create or select a Google Cloud project.
3. Enable **Google Sheets API** for that project.
4. In Google Auth Platform / the OAuth consent configuration:
   - complete the Branding details, including app name, support email, and
     developer contact;
   - choose the appropriate Internal or External audience;
   - add only the
     `https://www.googleapis.com/auth/spreadsheets.readonly` data-access scope
     for the current inspector;
   - add the user as a test user if the app remains in External testing mode.
5. Create an OAuth client ID with application type **Desktop app**.
6. Download the client JSON manually.
7. Create `%LOCALAPPDATA%\CursorExcelSkill\google\`.
8. Save the downloaded file as
   `%LOCALAPPDATA%\CursorExcelSkill\google\oauth-client.json`.
9. Install the official Python dependencies listed above.
10. Run `google_auth.py --authorize-readonly` for inspection-only work, or
    `google_auth.py --authorize-readwrite` for an explicitly requested write
    task, and complete browser consent.
11. Use `google_sheets_inspect.py` for independent read-only inspection and
    `google_sheets_write.py` for planned, verified writes.

Do not enable or use write scope unless the user explicitly requests a real
Google Sheets write operation.
