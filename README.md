# Cursor Skills

Reusable Agent Skills for Cursor and Cursor AI. This collection helps an AI
agent perform structured, repeatable work and synchronize the same skills
between machines, starting with production-oriented Excel, XLSX, Google Sheets,
and spreadsheet automation workflows.

## Available Skills

| Name | Purpose |
| --- | --- |
| [`/excel`](skills/excel/) | Create, edit, inspect, validate, and improve Excel XLSX and Google Sheets spreadsheets. |

The `/excel` skill covers:

- XLSX creation and editing with safe handling of existing spreadsheets
- Microsoft Excel Desktop recalculation through COM on Windows
- structural workbook inspection and validation
- rendered PDF-to-image visual QA
- formulas, Excel Tables, data validation, dropdowns, and checkboxes
- charts, dashboards, and financial or accounting workbooks
- official Google Sheets API integration with read-back verification

Platform-dependent capabilities are used only when their required software and
configuration are available.

## Installation

On Windows, clone the repository and run the installer from PowerShell:

```powershell
git clone https://github.com/Protside/cursor-skills.git
cd cursor-skills
.\install.ps1
```

The script discovers each directory under `skills/` and installs it into:

```text
%USERPROFILE%\.cursor\skills\
```

Cursor should then discover `/excel`. Runtime dependencies for specific
features may need to be installed separately.

## Updating

```powershell
git pull
.\install.ps1
```

## Usage

Describe the business task normally; the skill supplies the implementation,
safety, and QA workflow.

```text
/excel
Create a monthly expense tracker with a dashboard.
```

```text
/excel
Inspect this existing XLSX and add a payment-status column without breaking
existing formulas or formatting.
```

```text
/excel
Create a Google Sheets budget tracker with dropdowns, native checkboxes,
formulas, and charts.
```

## Excel / XLSX support

Python and `openpyxl` handle workbook structure, creation, editing, inspection,
and validation. On Windows, Microsoft Excel Desktop can provide authoritative
compatibility checks, formula recalculation, and PDF rendering through COM.
Those Excel-native QA capabilities require Windows, Microsoft Excel Desktop,
and `pywin32`.

## Google Sheets support

Google Sheets functionality uses the official Google Sheets API and OAuth 2.0.
Each user creates and configures their own OAuth Desktop credentials. Client
configuration and tokens are machine-local, are never included in this
repository, and must be configured separately on each computer.

## Dependencies

The core skill consists of Markdown instructions, references, and Python
scripts. Install runtime dependencies only for the capabilities you use:

- `openpyxl` — XLSX inspection, editing, and structural validation
- `pywin32` — Excel COM recalculation and PDF export on Windows
- `PyMuPDF` — rendering PDF pages to images for visual QA
- `google-api-python-client` — Google Sheets API operations
- `google-auth`, `google-auth-oauthlib`, and `google-auth-httplib2` — Google
  OAuth authentication and authorized API transport
- Microsoft Excel Desktop — Excel-native compatibility, recalculation, and
  rendering QA

## Repository Structure

```text
cursor-skills/
├── skills/
│   └── excel/
│       ├── SKILL.md
│       ├── references/
│       └── scripts/
├── install.ps1
└── WORKFLOW.md
```

## Creating New Skills

Create each canonical skill under `skills/<skill-name>/`, develop it in the
repository, and deploy it for local testing with `.\install.ps1`. The repository
copy is the source of truth; `%USERPROFILE%\.cursor\skills\` is only the
installed runtime copy. See [WORKFLOW.md](WORKFLOW.md) for the complete
repository-first development and release process.

## Security

Never commit OAuth tokens, OAuth client secrets, API keys, passwords, `.env`
secrets, credentials, private keys, or generated caches. Authentication and
machine-specific configuration remain local to each computer.

## License

This project is available under the [MIT License](LICENSE).

## Русский

Это репозиторий с переиспользуемыми Agent Skills для Cursor. Сейчас основной
навык — `/excel` для работы с Excel, XLSX и Google Sheets. Установка навыков
выполняется в PowerShell через `install.ps1`. В дальнейшем в коллекцию могут
быть добавлены новые навыки.
