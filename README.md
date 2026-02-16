# GO Contact Sync Mod JKFix

Private fork of GO Contact Sync Mod with a stable contacts-sync patch set for Outlook -> Google workflows.

## Why This Fork Exists

This fork exists to address recurring reliability issues seen by users on the original SourceForge-distributed build, especially for Outlook -> Google contacts-only sync.

Most common symptoms fixed here:

- repeated no-change contact updates on every sync run
- deletion prompts appearing even when **Sync Deletion** is OFF
- profile/folder selection not persisting reliably
- unclear startup behavior when Outlook is not running

## Keywords (Findability)

If you found this repo by searching, these are the issue phrases this fork targets:

- GO Contact Sync Mod repeated updates
- GO Contact Sync Mod sync deletion off still prompts
- GO Contact Sync Mod Loading Outlook folders stuck
- GO Contact Sync Mod Outlook to Google only re-sync bug
- GOContactSync duplicate warning cannot be synchronized

## Current Stable Release

- App informational version: `4.3.0-jkfix.4`
- Release tag: `v4.3.0-jkfix.4`
- Release page: `https://github.com/jaywking/go-contact-sync-mod-jkfix/releases/tag/v4.3.0-jkfix.4`

Direct download links:

- MSI installer: `https://github.com/jaywking/go-contact-sync-mod-jkfix/releases/download/v4.3.0-jkfix.4/SetupGCSM-4.3.0-jkfix.4.msi`
- ZIP (portable): `https://github.com/jaywking/go-contact-sync-mod-jkfix/releases/download/v4.3.0-jkfix.4/GOContactSyncMod-4.3.0-jkfix.4.zip`

## What This Fork Fixes

- Prevents repeated no-change Outlook -> Google contact updates on immediate re-sync.
- Ensures deletion prompts are suppressed when **Sync Deletion** is disabled.
- Persists and respects contact match IDs between runs.
- Improves Outlook folder loading behavior and startup usability.
- Adds Outlook pre-sync readiness checks (prompt or auto-start behavior).
- Fixes per-profile folder persistence so each sync profile reliably keeps its own selected source/target folders.
- Adds private build/install/package automation scripts.

## Quick Install (Pinned to jkfix.4)

1. Download MSI from the release page.
2. Run `SetupGCSM-4.3.0-jkfix.4.msi`.
3. Launch app and verify title shows `4.3.0-jkfix.4`.
4. Select correct sync profile and Outlook source folder.

Portable ZIP option:

1. Download `GOContactSyncMod-4.3.0-jkfix.4.zip`.
2. Extract anywhere.
3. Run `GOContactSync.exe`.

## Trust Signals / Validation

Before (problem state):

- Log repeatedly showed: `Updated contact from Outlook to Google: ...` for same unchanged contacts on every run.

After (jkfix expected behavior):

- First sync after profile/folder reset may update affected contacts.
- Immediate re-sync without changes should show:
  - `Sync complete. Synced: 0 ...`
  - unchanged contacts skipped as expected.
- Real field changes (for example phone edit in Outlook) should update only changed contact(s).

Known limitations:

- Contacts that match multiple Google contacts are intentionally skipped until duplicates are resolved.
- Upstream version checker may still reference upstream numeric line (`4.3.0`) while app title/log includes `jkfix` build.

SHA256 checksums for release assets:

- `GOContactSyncMod-4.3.0-jkfix.4.zip`  
  `5401B62116567DBAC45DD8FD819130BA18E8767F68C913FC888C8D1F97E1448A`
- `SetupGCSM-4.3.0-jkfix.4.msi`  
  `E6FCCACE46D87C0CDDB7BCD5AC0AF642E57D50D20BB938D1537A25AC05FFD5DA`

## Changelog

### 4.3.0-jkfix.4

- Installer upgrade reliability fix:
  - bumped numeric assembly/file version to `4.3.1` so MSI major upgrades replace binaries correctly
  - retained informational app version `4.3.0-jkfix.4` for user-facing build identity
- Prevents stale older `jkfix` binaries from remaining after MSI upgrade.

### 4.3.0-jkfix.3

- Added per-profile `Outlook start` setting in UI (`Prompt` / `Auto-start Outlook` / `Skip check`) with profile persistence.
- Added profile binding status line in UI to show active profile + selected Outlook contacts folder + Google account.
- Cleaned layout using designer-based control placement for the new settings.
- Improved sync log readability by bolding `Sync complete...` summary lines in the log pane.

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
- `VERIFICATION.md` - release verification checklist and proof template
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

- `dist\GOContactSyncMod-4.3.0-jkfix.4.zip`

## Known Notes

- Duplicate-contact warnings are expected when Outlook contacts map to multiple Google contacts; those entries are skipped until duplicates are resolved.
- Installer (`SetupGCSM`) build may fail without WiX installed; EXE-only build remains valid for private usage.

## Upstream

- Original project source: SourceForge SVN/Git mirror for GO Contact Sync Mod.
