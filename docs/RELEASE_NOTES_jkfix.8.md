# Release 4.3.0-jkfix.8

Display version **4.3.0-jkfix.8**, tag **v4.3.0-jkfix.8**, numeric assembly/file/installer version **4.3.5**. The MSI upgrades jkfix.7 (numeric 4.3.4) using the existing remembered-install-path policy. MSI and portable ZIP packages include the support, release, and testing documentation.

## Included fixes

1. **Preserve Google favorites.** In Outlook to Google Only mode, contact updates retain Google's starred group even when Outlook has no matching category. Ordinary Google labels still follow Outlook categories; two-way category behavior is unchanged.
2. **Detect edits made immediately after a sync.** Linked contacts in Outlook to Google Only mode compare mapped content instead of relying solely on the 120-second timestamp tolerance. Real edits remain detectable; timestamp-only bookkeeping changes do not trigger repeat writes. Other sync modes and calendar change detection retain their existing behavior.
3. **Report failed saves accurately.** False save results and final write exceptions count as errors once in the contact save loop. A failed contact cannot subtract from the successful-save count. Logs retain the Google API reason after retries; a recovered retry counts as success.
4. **Isolate contact preparation failures.** An unreadable contact, photo, or baseline is reported and excluded from the later save/delete phase for that run. Other contacts continue. Existing baselines are retained for retry, and failed preparation cannot leave replacement matches queued for deletion. Cancellation still stops the sync.
5. **Use a green tray icon.** The existing green application icon replaces the legacy grayscale tray artwork, including all rotating sync-animation frames. The separate warning/error indicator remains distinct.
6. **Compact the Automation section.** Its three checkbox rows use the same vertical pitch as the Sync Options list. Interval controls and the countdown fit within a shorter group, with the recovered space available to Sync Options. The countdown row stays reserved when hidden to prevent layout jumps.
7. **Remember window placement.** Size, position, and maximized state survive tray reopening and restarts. First-run sizing fits the monitor, and restored bounds are constrained to an available monitor's working area. Normal bounds remain intact while minimized or maximized. Placement is per Windows user, independently of sync profiles.

## Upgrade behavior and stored settings

The first content-tracking run records an observed baseline for existing contacts that legacy checks consider unchanged, without rewriting Google. Already-due updates and phone repairs still follow their normal paths. Older missed edits are not automatically reconciled: make a fresh edit after initialization to send them.

Content baselines contain hashes rather than contact field values under `%APPDATA%\GoContactSyncMOD\ContentBaselines\v1`. They are isolated by profile, Google account, source folder, selected Google label, and contact. Failed tracked writes remain pending across restarts. A successful baseline requires the contact, enabled photo operations, and Outlook link metadata to save successfully.

Corrupt baselines are retained and may continue to fail until repaired. Do not reset all baselines or matches to work around one failing contact. Isolation covers the per-contact preparation/save stages, not global connection, loading, matching, or group-sync failures, and does not roll back writes already made during preparation in other sync modes. Ambiguous duplicate matches are still skipped for manual resolution.

Window placement is stored as the `WindowPlacement` string value under `HKCU\Software\GoContactSyncMOD`. Missing, malformed, or unreadable saved placement falls back to screen-fitting defaults. A persistence failure is logged without stopping sync.

## Validation

- The integrated `4.3.0-jkfix.8` Release build passed **66 automated tests**: 10 favorites, 18 content tracking, 16 save reporting, 9 preparation isolation, 12 window placement, and 1 same-email test.
- The user confirmed successive company edits reached Google while the favorite star remained, and completed the earlier no-change/restart checks. The user also reported improved window behavior after trying the window preview. This does not claim every manual scenario in the testing guide has been completed.
- Tray frames were visually checked at 16px and 32px on light/dark backgrounds. Native Automation layout checks confirmed equal row pitches at 10pt/15pt fonts, no clipped controls, and stable repeated layout.
- Failure tests use simulated services, isolated forms, or temporary files. Production-contact failure injection was not performed.
- Release application and installer builds passed. MSI metadata and extracted binaries are checked against the previous release and the portable ZIP. A live uninstall/reinstall or in-place upgrade was not performed on the production workstation.

See [fix details](CONTACT_SYNC_FIXES.md), [test instructions](TESTING_CONTENT_TRACKING.md), and [support guidance](../SUPPORT.md).
