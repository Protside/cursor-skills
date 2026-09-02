# Cursor Skills

Private, Git-backed storage for personal Cursor Agent Skills. This repository is
the canonical source for versioning, backing up, and synchronizing skills across
Windows computers.

## Source-of-truth rule

**Repository first; global is deployment.**

- `skills/<skill-name>/` is the source, development workspace, and
  version-controlled copy.
- `$env:USERPROFILE\.cursor\skills\` is an installed runtime copy used for local
  testing. Do not make permanent edits there.
- `origin/main` is the backed-up, synchronized, approved collection of tested
  skills.

All permanent changes must begin in this repository. Deploy them locally with
`.\install.ps1`, test the installed copy in Cursor, and make any fixes back in
the repository before reinstalling and retesting.

## Structure

```text
cursor-skills/
├── README.md
├── WORKFLOW.md
├── .gitignore
├── install.ps1
└── skills/
    └── excel/
```

Each immediate subdirectory of `skills/` is one installable Cursor skill.

## Add another skill

1. Create `skills/<skill-name>/`.
2. Add `SKILL.md` and all non-secret `references/`, `scripts/`, `assets/`,
   `evals/`, and other resources beneath that directory.
3. Develop only in the repository copy.
4. Run `.\install.ps1`, then test the installed skill in a relevant Cursor
   project and fresh chat.
5. Fix problems in the repository, reinstall, and retest until the skill passes.
6. Review Git status and secret exclusions before committing and pushing.

See [WORKFLOW.md](WORKFLOW.md) for the complete lifecycle.

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

For an existing skill, edit its directory under `skills/`, run the installer,
test the global runtime copy, and repeat until it passes.

## Publish changes

Publish only stable, tested repository changes:

```powershell
git add .
git commit -m "Describe the skill changes"
git push origin main
```

## Credentials and tokens

OAuth credentials, access tokens, refresh tokens, client secrets, environment
secrets, and other authentication data are intentionally machine-local and are
never stored in Git. This also excludes API keys, passwords, `.env` secrets,
credential directories, `__pycache__/`, and `*.pyc`.

In particular, Google authentication data for the Excel skill remains under
`%LOCALAPPDATA%\CursorExcelSkill\google\` and is not part of this repository.
Never manually synchronize global skill folders between computers when this Git
repository is available; use `git pull` followed by `.\install.ps1`.
