# Contact sync fixes in 4.3.0-jkfix.8

This release integrates all fixes in one source tree. Display version **4.3.0-jkfix.8**, numeric version **4.3.5**, tag **v4.3.0-jkfix.8**. See [release notes](RELEASE_NOTES_jkfix.8.md) and the [download links](../README.md#current-stable-release).

The work comprises four sync fixes and three interface improvements:

| Change | Behavior and scope |
| --- | --- |
| Favorites preservation | In **Outlook to Google Only**, updating a contact retains membership in Google's `contactGroups/starred` group even without an Outlook category. Ordinary labels still follow Outlook categories. Two-way category behavior is unchanged. |
| Content-based change detection | Linked contacts in **Outlook to Google Only** detect changes to mapped content even within the former 120-second tolerance. Timestamp-only bookkeeping does not cause repeat updates. Other modes and calendar change detection retain their existing behavior. |
| Failed-save reporting | A false save result or final write exception counts as one error in the contact save loop. Failures do not subtract from successful saves. Google API errors retain their reason after the configured retries; a recovered retry counts as success. |
| Preparation failure isolation | A contact that fails preparation is reported once and excluded from the later save/delete phase. Other contacts continue. Existing baselines are retained, and preparation is attempted on a later sync. Cancellation still stops the run. |
| Green tray icon | The normal tray icon and rotating sync frames use the green application artwork. The warning/error indicator remains distinct. |
| Compact Automation layout | Checkbox spacing matches the Sync Options list; interval/countdown controls fit in a shorter group without a height jump when Auto Sync is toggled. |
| Remembered window placement | Size, position, and maximized state are restored across tray reopening and restarts, with screen-fitting defaults and bounds recovery when monitors change. |

## Baselines and recovery

On the first run, existing contacts considered unchanged by the legacy checks receive an observed baseline without rewriting Google. This is the chosen migration policy. Older missed edits are not automatically reconciled: make a fresh edit after initialization. Already-due updates and existing phone-repair behavior can still run.

Baselines contain hashes, not contact field values, under `%APPDATA%\GoContactSyncMOD\ContentBaselines\v1`. They are isolated by profile, Google account, Outlook folder, selected Google label, and contact. Failed tracked writes remain pending across restarts. A completed baseline is recorded only after the contact, Outlook link metadata, and enabled photo operations succeed.

A corrupt baseline is retained and reported, not silently replaced. It may continue to fail until its cause is repaired. Do not clear all baselines to resolve one contact: that would reset change tracking and could hide previously missed edits. Preparation isolation does not provide rollback of writes already made during preparation in other sync modes, and does not isolate global connection, loading, matching, or group-sync failures.

## Interface behavior

The app now remembers normal window bounds and maximized state across tray reopening and restarts, instead of forcing a default normal window. First-run bounds are centered and limited to the monitor's working area; saved bounds are fitted to an available monitor if the display setup changes. Normal bounds are retained while minimized or maximized. Placement is stored in the `WindowPlacement` string value under `HKCU\Software\GoContactSyncMOD`, independently of sync profiles. Malformed or unreadable placement falls back to the default, and persistence failures are logged without stopping sync.

The Automation section now uses the Sync Options checkbox row pitch and fits its interval/countdown controls, returning excess height to Sync Options. Native control rendering and geometry checks confirmed matching 20px/29px pitches at 10pt/15pt fonts, no clipped controls, and stable height on repeated layout.

This build also replaces the legacy gray tray artwork with the existing green application icon, including rotated sync-animation frames. The warning/error icon remains distinct. The frames were visually checked at 16px and 32px on light and dark backgrounds.

## Running and building this release

Exit the existing app, including its tray process, before installing the MSI or launching `GOContactSync.exe` from the extracted portable ZIP. Both use your existing profiles and contact data. Verify the displayed version is `4.3.0-jkfix.8`.

All source and fixtures are included in tag `v4.3.0-jkfix.8`; no temporary patch or second checkout is required. Build `GoogleContactsSync/GoogleContactsSync.csproj` and `UnitTests/UnitTests.csproj` in Release/x86 with restored NuGet packages and .NET Framework 4.8.1 reference assemblies. Installer packaging uses `SetupGCSM/SetupGCSM.wixproj` with WiX 3.14.1. Use the fixture filter in the testing guide to avoid live-contact legacy tests.

## Validation

The integrated release build passed **66 automated tests**: 12 window placement, 10 favorites, 18 content tracking, 16 save reporting, 9 preparation isolation, and 1 same-email test. Failure tests use simulated services or temporary local files; they do not access live Google or Outlook contacts.

The earlier combined content/favorites preview passed the user's live checks: both successive company edits reached Google, the star remained, and no-change/restart checks were completed. The user reported improved window behavior after trying the window preview. Failed-save reporting and preparation isolation have automated coverage; their live failure paths have not been deliberately exercised. Not every manual scenario in the testing guide has been confirmed.

See [the testing guide](TESTING_CONTENT_TRACKING.md) for the fixture filter and live checks, and [support guidance](../SUPPORT.md) for reporting failures.
