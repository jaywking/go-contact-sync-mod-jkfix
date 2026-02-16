# GO Contact Sync Mod JKFix

Private fork of GO Contact Sync Mod with a stable contacts-sync patch set for Outlook -> Google workflows.

## Version

- App informational version: `4.3.0-jkfix.2`
- Release tag: `v4.3.0-jkfix.2`

## What This Fork Fixes

- Prevents repeated no-change Outlook -> Google contact updates on immediate re-sync.
- Ensures deletion prompts are suppressed when **Sync Deletion** is disabled.
- Persists and respects contact match IDs between runs.
- Improves Outlook folder loading behavior and startup usability.
- Adds Outlook pre-sync readiness checks (prompt or auto-start behavior).
- Fixes per-profile folder persistence so each sync profile reliably keeps its own selected source/target folders.
- Adds private build/install/package automation scripts.

## Changelog

### 4.3.0-jkfix.2

- Added Outlook readiness pre-check before Sync/Reset:
  - default prompt mode when Outlook is not running
  - optional auto-start mode via environment variable
- Fixed per-profile folder persistence so each profile keeps its own selected Outlook/Google folders after restart.
- Added/updated installer and release packaging flow for MSI + ZIP publishing.

### 4.3.0-jkfix.1

- Fixed repeated no-change Outlook -> Google re-updates on immediate reruns.
- Suppressed deletion prompts when Sync Deletion is off.
- Improved match ID persistence/stability across sync runs.
- Added startup/folder scan usability improvements and private build tooling.

## Repository Layout

- `GoogleContactsSync/` - main application source
- `scripts/build-private.ps1` - restore + build script
- `scripts/install-private.ps1` - local installer-style deployment (Start Menu/Desktop shortcut)
- `scripts/package-private.ps1` - creates release ZIP
- `patches/contact-resync-fix-r1587.patch` - patch export
- `TEST_PLAN.md` - regression checks
- `NEXT_STEPS.md` - operational notes

## Build

From repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-private.ps1 -Configuration Release -Platform x86
```

Output EXE:

- `GoogleContactsSync\bin\Release\GOContactSync.exe`

## Install Like a Regular App

Current-user install:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-private.ps1 -Scope CurrentUser
```

All-users install (requires elevated PowerShell):

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-private.ps1 -Scope AllUsers
```

Uninstall:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-private.ps1 -Scope CurrentUser -Uninstall
```

## Package ZIP

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\package-private.ps1
```

Output:

- `dist\GOContactSyncMod-4.3.0-jkfix.2.zip`

## Known Notes

- Duplicate-contact warnings are expected when Outlook contacts map to multiple Google contacts; those entries are skipped until duplicates are resolved.
- Installer (`SetupGCSM`) build may fail without WiX installed; EXE-only build remains valid for private usage.

## Upstream

- Original project source: SourceForge SVN/Git mirror for GO Contact Sync Mod.
