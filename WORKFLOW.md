# Cursor Skill Development and Release Workflow

## Governing rule

**REPOSITORY-FIRST, GLOBAL-IS-DEPLOYMENT.**

- `skills/<skill-name>/`: source, development, and version control.
- `%USERPROFILE%\.cursor\skills\`: installed runtime copy for testing.
- `origin/main`: backup, synchronization, and latest approved version.

The repository is the canonical source of truth for every personal Cursor
skill. Never develop in or permanently patch the global skills directory. If a
runtime test reveals a problem, fix the repository copy and reinstall it.

## Complete lifecycle

```text
CREATE
  -> DEVELOP IN REPOSITORY
  -> INSTALL LOCALLY
  -> TEST
  -> FIX IN REPOSITORY
  -> REINSTALL
  -> RETEST
  -> PASS
  -> COMMIT
  -> PUSH
  -> PULL + INSTALL ON OTHER DEVICES
```

## Create a skill

1. Create `skills/<skill-name>/`.
2. Put all non-secret skill files beneath it, including `SKILL.md` and any
   `references/`, `scripts/`, `assets/`, `evals/`, or other resources.
3. Develop and edit only this repository copy.
4. Install the current repository version:

   ```powershell
   .\install.ps1
   ```

5. Test the installed skill in Cursor using a fresh or relevant project and
   chat.
6. If testing fails, fix `skills/<skill-name>/`, reinstall, and retest.
7. Continue until functional tests and QA pass.
8. Inspect Git status and confirm that no secrets or generated caches will be
   tracked.
9. Commit and push the stable version to `origin/main`.

## Update an existing skill

For example, to update `/excel`:

1. Edit `skills/excel/`.
2. Install:

   ```powershell
   .\install.ps1
   ```

3. Test the global installed copy.
4. Fix all problems in `skills/excel/`, never in the global copy.
5. Reinstall and retest until the result passes.
6. Commit and push the approved repository changes.

## Release approved changes

```powershell
git status
git add .
git commit -m "Describe the tested skill changes"
git push origin main
```

Before committing, verify that staged files contain no credentials, tokens,
secrets, machine-local configuration, or generated caches.

## Set up another Windows computer

Clone once:

```powershell
git clone https://github.com/Protside/cursor-skills.git
cd cursor-skills
.\install.ps1
```

For later updates:

```powershell
git pull
.\install.ps1
```

Do not manually copy global skill directories between computers when the
repository is available.

## Security boundary

Never commit OAuth tokens, OAuth client secrets, API keys, passwords, `.env`
secrets, machine-local authentication files, credential directories,
`__pycache__/`, or `*.pyc`.

Authentication and machine-specific configuration remain local to each
computer. Excel skill Google credentials remain outside Git at:

```text
%LOCALAPPDATA%\CursorExcelSkill\google\
```

Software dependencies such as Python packages may need separate installation on
each computer.
