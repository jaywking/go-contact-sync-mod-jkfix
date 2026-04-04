# Outlook Troubleshooting

Use this guide when `jkfix` cannot connect to Outlook or reports Office registry/configuration problems.

## Common Symptoms

- `Could not connect to Microsoft Outlook`
- Outlook folder loading never completes
- Errors mentioning Office registry keys, `TypeLib`, or Click-to-Run
- Outlook appears installed, but the app cannot detect it correctly

## Quick Checks

1. Close both Outlook and GO Contact Sync Mod, then reopen Outlook first.
2. Confirm Outlook opens normally with your expected profile.
3. Confirm Outlook bitness and installed Office version.
4. Re-run GO Contact Sync Mod after Outlook is fully loaded.

## Most Common Causes

- Broken Office COM registration
- Old Office registry entries left behind after upgrade/uninstall
- Click-to-Run and MSI Office remnants conflicting
- Outlook installed in a different bitness/configuration than expected

## What To Try

1. Run an Office repair from Windows Apps / Installed Apps.
2. Reboot after the repair.
3. If you recently changed Office versions, remove leftover older Office installs.
4. Reinstall Outlook/Office if COM registration remains broken after repair.
5. If possible, avoid mixed or partially removed Office installations on the same machine.

## Before Opening An Issue

Collect:

- App version from the window title or About dialog
- Outlook version and bitness
- The full error message
- Relevant log lines from:
  `C:\Users\<your-user>\AppData\Roaming\GoContactSyncMOD\`

Then open an issue:

- `https://github.com/jaywking/go-contact-sync-mod-jkfix/issues/new/choose`
