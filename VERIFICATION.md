# Verification Guide

Use this file to validate each `jkfix` release before publishing and to capture proof artifacts for users.

## Scope

- Contacts-only sync path
- One-way mode: Outlook -> Google
- Profile persistence and startup behavior
- Release asset integrity

## Baseline Config

- `Sync Contacts`: ON
- `Sync Appointments`: OFF
- `Sync Deletion`: OFF
- `Merge` options: OFF
- Sync direction: `Outlook To Google Only`

## Functional Verification

1. Immediate re-sync should not churn unchanged contacts
- Run sync once.
- Run sync again without edits.
- Expected: `Sync complete. Synced: 0 ...` (or no repeated unchanged-contact writes).

2. Real edit should sync exactly changed data
- Change one known field in Outlook (example: mobile phone).
- Run sync.
- Expected: only changed contact(s) update.
- Run sync again with no edits.
- Expected: no repeat write for unchanged contact.

3. Deletion prompt behavior when deletion is OFF
- Keep `Sync Deletion` OFF.
- Run sync.
- Expected: no deletion confirmation dialog.

4. Per-profile persistence
- In profile A: choose Outlook folder + startup mode.
- In profile B: choose different folder + startup mode.
- Restart app and switch profiles.
- Expected: each profile restores its own values.

5. Outlook startup mode behavior
- `Prompt`: asks when Outlook is closed.
- `Auto-start Outlook`: launches Outlook and continues.
- `Skip check`: no prompt before sync.

## Log Evidence to Capture

Capture from:
- `C:\Users\<your-user>\AppData\Roaming\GoContactSyncMOD\`

Keep snippets for release notes:
- one successful no-change immediate re-sync
- one real-change sync (single contact)
- one startup-mode behavior sample

## Release Asset Verification

Generate SHA256 hashes:

```powershell
Get-FileHash .\dist\GOContactSyncMod-<version>.zip -Algorithm SHA256
Get-FileHash .\dist\SetupGCSM-<version>.msi -Algorithm SHA256
```

Publish both hashes in:
- `README.md`
- GitHub Release notes

## Version Confirmation

Confirm build identity in app:
- Window title includes `4.3.0-jkfix.x`
- Startup log includes installed app version

## Sign-off Template

Use this block before release:

```text
Release: v4.3.0-jkfix.x
Date:
Tester:
No-change immediate resync: PASS/FAIL
Real field edit propagation: PASS/FAIL
Deletion prompt suppression (deletion OFF): PASS/FAIL
Per-profile folder persistence: PASS/FAIL
Outlook startup mode behavior: PASS/FAIL
ZIP hash published: YES/NO
MSI hash published: YES/NO
Notes:
```
