# Support

This repository is the support home for the `jkfix` fork of GO Contact Sync Mod.

The [jkfix.8 release notes](docs/RELEASE_NOTES_jkfix.8.md) describe all four sync fixes and three interface improvements in version `4.3.0-jkfix.8` (numeric installer version `4.3.5`).

## Where To Report Issues

- Search existing issues first: `https://github.com/jaywking/go-contact-sync-mod-jkfix/issues`
- Open a new issue here: `https://github.com/jaywking/go-contact-sync-mod-jkfix/issues/new/choose`

## What To Include

- Installed app version shown in the window title or About dialog
- Whether you installed via MSI, ZIP, or a local preview; for previews include the executable path and `BUILD-INFO.json` version
- Outlook version and bitness
- Whether the sync is contacts, appointments, or both
- Exact steps to reproduce
- Relevant log excerpt from:
  `C:\Users\<your-user>\AppData\Roaming\GoContactSyncMOD\`

## In Scope

- Regressions in `jkfix` releases
- Outlook to Google contacts sync issues
- Installer/update problems for `jkfix` MSI releases
- Startup, profile persistence, and folder-selection issues introduced by the fork

## Out Of Scope

- Support for the upstream SourceForge release line
- Microsoft Outlook installation repair beyond basic troubleshooting guidance
- Google account or Google Workspace administrative issues outside app behavior

## Outlook Connection Problems

If the app cannot connect to Outlook, start here:

- `docs/TROUBLESHOOTING_OUTLOOK.md`

## Contact Errors

See [fix details](docs/CONTACT_SYNC_FIXES.md) for scope and recovery behavior and [testing instructions](docs/TESTING_CONTENT_TRACKING.md) for expected behavior.

- `Failed to prepare contact`: that contact was excluded from the later save/delete phase. Other contacts can continue. Check the underlying read, photo, or baseline error before retrying.
- `Failed to synchronize contact`: saving failed. Include the Google/API reason and final sync summary; do not assume the cause is an oversized notes field.
- A corrupt baseline is retained and may fail again until repaired. Preserve the affected file and error details for diagnosis; do not delete the baseline directory or reset all matches as a routine remedy.

Logs are under `%APPDATA%\GoContactSyncMOD`. Review shared excerpts for personal contact details and never include OAuth token files.

## Window, Layout, And Version Messages

- Include display scaling, monitor changes, whether the app was reopened from the tray or restarted, and whether the window was normal or maximized. Window placement is stored separately from sync profiles under `HKCU\Software\GoContactSyncMOD` in the `WindowPlacement` value.
- Automation rows should match the Sync Options row spacing. Report clipped controls or unexpected height changes when toggling Auto Sync.
- The tray is green during ordinary use and sync animation; an error uses the distinct warning indicator.
- The update check compares numeric installer versions: jkfix.8 uses `4.3.5`, and jkfix.7 used `4.3.4`. The installed build text identifies the exact informational version. A local preview may share an installer number with a release; a numeric comparison alone does not establish which fixes it contains.
