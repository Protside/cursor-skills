# Cursor Skills

Private, Git-backed storage for personal Cursor Agent Skills. This repository is
the canonical source for versioning, backing up, and synchronizing skills across
Windows computers.

## Structure

```text
cursor-skills/
├── README.md
├── .gitignore
├── install.ps1
└── skills/
    └── excel/
```

Each immediate subdirectory of `skills/` is one installable Cursor skill.

## Add another skill

1. Create or copy the complete skill directory into `skills/<skill-name>/`.
2. Confirm that the skill contains its `SKILL.md` and any required supporting
   files.
3. Keep credentials, tokens, secrets, and machine-local authentication data out
   of the repository.
4. Review the changes with `git status` before committing.

## Install on a new Windows PC

1. Install Git and Cursor.
2. Clone this private repository.
3. Open PowerShell in the repository root.
4. Run:

   ```powershell
   .\install.ps1
   ```

The installer discovers every directory under `skills/` and safely updates the
corresponding directory under `$env:USERPROFILE\.cursor\skills\`.

Software dependencies used by individual skills, such as Python and Python
packages, may need to be installed separately on each new computer.

## Update installed skills

```powershell
git pull
.\install.ps1
```

## Publish changes

```powershell
git add .
git commit -m "Describe the skill changes"
git push
```

## Credentials and tokens

OAuth credentials, access tokens, refresh tokens, client secrets, environment
secrets, and other authentication data are intentionally machine-local and are
never stored in Git. In particular, Google authentication data for the Excel
skill remains under `%LOCALAPPDATA%\CursorExcelSkill\google\` and is not part of
this repository.
