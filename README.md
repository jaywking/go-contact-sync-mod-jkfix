# GO Contact Sync Mod JKFix

Private fork of GO Contact Sync Mod with a stable contacts-sync patch set for Outlook -> Google workflows.

## Version

- App informational version: `4.3.0-jkfix.1`
- Release tag: `v4.3.0-jkfix.1`

## What This Fork Fixes

- Prevents repeated no-change Outlook -> Google contact updates on immediate re-sync.
- Ensures deletion prompts are suppressed when **Sync Deletion** is disabled.
- Persists and respects contact match IDs between runs.
- Improves Outlook folder loading behavior and startup usability.
- Adds private build/install/package automation scripts.

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

- `dist\GOContactSyncMod-4.3.0-jkfix.1.zip`

## Known Notes

- Duplicate-contact warnings are expected when Outlook contacts map to multiple Google contacts; those entries are skipped until duplicates are resolved.
- Installer (`SetupGCSM`) build may fail without WiX installed; EXE-only build remains valid for private usage.

## Upstream

- Original project source: SourceForge SVN/Git mirror for GO Contact Sync Mod.

