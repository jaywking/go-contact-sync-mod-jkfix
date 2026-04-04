# GO Contact Sync Mod JKFix

Unofficial maintained fork of GO Contact Sync Mod with a stable contacts-sync patch set for Outlook -> Google workflows.

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

- App informational version: `4.3.0-jkfix.6`
- Release tag: `v4.3.0-jkfix.6`
- Release page: `https://github.com/jaywking/go-contact-sync-mod-jkfix/releases/tag/v4.3.0-jkfix.6`

Direct download links:

- MSI installer: `https://github.com/jaywking/go-contact-sync-mod-jkfix/releases/download/v4.3.0-jkfix.6/SetupGCSM-4.3.0-jkfix.6.msi`
- ZIP (portable): `https://github.com/jaywking/go-contact-sync-mod-jkfix/releases/download/v4.3.0-jkfix.6/GOContactSyncMod-4.3.0-jkfix.6.zip`

## Support

- Download updates from GitHub Releases: `https://github.com/jaywking/go-contact-sync-mod-jkfix/releases`
- Report bugs and regressions on GitHub Issues: `https://github.com/jaywking/go-contact-sync-mod-jkfix/issues`
- Support guide: `SUPPORT.md`
- Outlook/Office troubleshooting: `docs/TROUBLESHOOTING_OUTLOOK.md`

## What This Fork Fixes

- Prevents repeated no-change Outlook -> Google contact updates on immediate re-sync.
- Ensures deletion prompts are suppressed when **Sync Deletion** is disabled.
- Persists and respects contact match IDs between runs.
- Improves Outlook folder loading behavior and startup usability.
- Adds Outlook pre-sync readiness checks (prompt or auto-start behavior).
- Fixes per-profile folder persistence so each sync profile reliably keeps its own selected source/target folders.
- Adds private build/install/package automation scripts.

## Quick Install (Pinned to jkfix.6)

1. Download MSI from the release page.
2. Run `SetupGCSM-4.3.0-jkfix.6.msi`.
3. Launch app and verify title shows `4.3.0-jkfix.6`.
4. Select correct sync profile and Outlook source folder.

Portable ZIP option:

1. Download `GOContactSyncMod-4.3.0-jkfix.6.zip`.
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

SHA256 checksums for release assets:

- `GOContactSyncMod-4.3.0-jkfix.6.zip`  
  `A76DE14F19B212CCD1C0512A6E27CC4C96B2B22D8D068D0E89DD5C256C617D5E`
- `SetupGCSM-4.3.0-jkfix.6.msi`  
  `DE3C0D34306C0D8562AB4BBCFECBBB7EF8DE07A5250EF2D803604C89F939C1B8`

## Changelog

### 4.3.0-jkfix.6

- Added a confirmation dialog before revoking saved Google authorization.
- Redesigned the profile binding summary into readable multi-line labels.
- Updated the main settings window to use `Segoe UI 10pt`.
- Reflowed the left-side settings layout to use space more cleanly when appointments are disabled.
- Bumped numeric assembly/file version to `4.3.3` so the installed MSI upgrades in place over `4.3.2`.

### 4.3.0-jkfix.5

- Repointed in-app update, support, homepage, and issue-reporting links to GitHub.
- Added fork-owned support and Outlook troubleshooting documentation.
- Fixed version display so UI surfaces prefer the `jkfix` informational version.
- Bumped numeric assembly/file version to `4.3.2` so MSI upgrades replace existing `4.3.1` installs.

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

## Updating an Existing JKFix Install

1. Download the newest `SetupGCSM-<version>.msi` from GitHub Releases.
2. Run the MSI over your existing installed `jkfix` build.
3. Launch the app from Start Menu and confirm the window title shows the new `jkfix` version.

Do not install upstream SourceForge releases over this fork unless you intentionally want to switch away from `jkfix`.

## Package ZIP

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\package-private.ps1
```

Output:

- `dist\GOContactSyncMod-4.3.0-jkfix.6.zip`

## Known Notes

- Duplicate-contact warnings are expected when Outlook contacts map to multiple Google contacts; those entries are skipped until duplicates are resolved.
- Installer (`SetupGCSM`) build may fail without WiX installed; EXE-only build remains valid for private usage.

## Upstream

- Original project source: SourceForge SVN/Git mirror for GO Contact Sync Mod.
