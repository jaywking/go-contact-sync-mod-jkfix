# Contact sync validation: 4.3.0-jkfix.8

Content tracking applies to linked contacts in **Outlook to Google Only** mode. Calendar and two-way change detection keep their existing behavior.

## Automated checks

Build `UnitTests/UnitTests.csproj` with the repository's restored packages and .NET Framework 4.8.1 reference assemblies. All fixtures are included in the release source tree; see [build guidance](CONTACT_SYNC_FIXES.md#running-and-building-this-release). Run only the following NUnit fixtures using VSTest:

```text
/TestCaseFilter:FullyQualifiedName~WindowPlacementTests|FullyQualifiedName~ContactContentTrackingTests|FullyQualifiedName~GoogleGroupPreservationTests|FullyQualifiedName~ContactSaveReportingTests|FullyQualifiedName~ContactPreparationIsolationTests|FullyQualifiedName~SameEmailTests
```

Do not run the entire legacy suite as a local unit check: other fixtures log into Outlook and Google and can create/delete test records. The integrated Release build passed all 66 tests.

The content tests cover edits inside the former two-minute window, unchanged contents with changed timestamps, first-run observation, account isolation, relinking, restart recovery, unsuccessful writes, edits during a request, corrupt/unwritable local state, and fingerprint stability for metadata, categories, notes, photos, and mapping settings.

Save-reporting tests cover false returns, exceptions, continuation to later contacts, preserved Google error details, successful retries, exhausted retries, and accurate counters. Preparation tests cover partially mapped contacts, deletion prevention, corrupt and pending baselines, retry after recovery, cancellation, failed conflict replacements, and a missing cached Outlook ID.

The MSI and ZIP contain all four sync fixes and three interface improvements listed in the [jkfix.8 release notes](RELEASE_NOTES_jkfix.8.md). The display version is `4.3.0-jkfix.8`; numeric assembly/file/installer version is `4.3.5`.

## Live test sequence

1. Exit the currently running GO Contact Sync Mod, including its tray process. Launch the installed app or `GOContactSync.exe` from the extracted release ZIP and check its displayed version. Confirm that the tray icon is green and stays green while rotating during a normal sync. An error still shows the separate warning indicator.
2. Select **Outlook to Google Only**, confirm the source folder and Google account, and run one sync to initialize tracking. Existing contacts that the legacy checks consider unchanged receive an observed local baseline. Already-due updates and existing phone-repair behavior can still run normally.
3. Choose a linked test contact. Change its company in Outlook, save and close the contact, then sync. Confirm Google shows the new company.
4. Within two minutes of that sync, change the company again, save and close, and sync immediately. Confirm the second value reaches Google and its Google favorite star remains as well.
5. Run another sync without edits. The test contact should not be updated. Restart the app and repeat the no-change check.
6. Restore the test contact's original company and sync; this edit should also be detected immediately.

The app uses real profile settings and contact data. Baselines contain hashes rather than field values and are stored under `%APPDATA%\GoContactSyncMOD\ContentBaselines\v1`. First-run observation deliberately does not reconcile previously missed edits; a fresh edit after initialization is required for those contacts.

An unsuccessful update is kept pending until the contact, enabled photo operations, and Outlook link metadata have saved successfully. The baseline uses the contents captured before the outgoing update, so an edit made during a save remains detectable on the next run.

## Checking failure behavior

For window placement, resize and move the app, hide it to the tray, then reopen it. Fully exit and restart it and confirm the same bounds return. Repeat while maximized, including minimizing to the tray; it should reopen maximized and retain the earlier normal size when restored. With a monitor disconnected or a smaller working area, the restored window should remain on-screen. Tests cover these transitions using isolated forms and injected storage without changing your real window settings.

For the compact Automation layout, confirm its three checkbox rows have the same spacing as the Sync Options list. Toggle Auto Sync to show/hide the countdown and resize the window; the interval controls should remain readable and the group should not jump in height.

Use the automated fixtures to induce errors; do not corrupt real baselines or contacts for testing. If a real failure occurs during normal use:

- A preparation failure logs `Failed to prepare contact`, identifies the affected contact, and excludes it from the later save/delete phase. Other contacts continue, and the error is counted once.
- A save failure logs `Failed to synchronize contact` with the underlying reason. Successful-save counts remain accurate; an exhausted retry counts as an error, while a recovered retry counts as success.
- A tracked pending update should retry after the underlying issue is fixed. Corrupt baselines remain errors until repaired; waiting alone does not repair them.
- Cancellation should stop the sync rather than become an ordinary per-contact error.

Record the displayed version, final summary, and nearby error lines. Follow [support guidance](../SUPPORT.md). No-change syncs validate ordinary behavior but do not exercise the failure paths.
