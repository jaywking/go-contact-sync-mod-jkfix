## Regression Test Plan (Contacts-Only, One-Way Outlook -> Google)

### Profile Under Test
- Sync option: `OutlookToGoogleOnly`
- Sync deletion: `OFF`
- Merge options: `OFF`
- Contacts sync: `ON`
- Appointments sync: `OFF`

### Pre-check
1. Export/capture baseline logs.
2. Note count of contacts expected to sync.
3. Confirm 26 previously problematic contacts still exist on both sides.

### Test 1: No Delete Dialogs With Sync Deletion Off
1. Run sync.
2. Verify no deletion prompt appears:
- No `Outlook Contact deleted` dialog.
- No `Google Person deleted` dialog.
3. Check logs for:
- `Skipped deletion prompt ... SyncDeletion is switched off`

### Test 2: No Re-update Loop For Unchanged Contacts
1. Run sync once.
2. Run sync again without any edits.
3. Verify second run does not repeatedly log:
- `Updated contact from Outlook to Google` for unchanged contacts.
4. For previously noisy entries, confirm skip log:
- `Skipping Outlook-to-Google overwrite because only timestamp changed and etag is unchanged`

### Test 3: Match IDs Persist Between Runs
1. Run sync.
2. Restart app.
3. Run sync again.
4. Verify matching is ID-based and stable (no re-churn).
5. In logs, allow one-time:
- `Persisted contact match IDs for: "..."`

### Test 4: Real Change Still Syncs
1. Edit one Outlook contact field (phone or title).
2. Run sync.
3. Verify exactly that contact updates on Google.
4. Run sync again with no changes; verify no repeat update.

### Test 5: Controlled Deletion Behavior Check
1. Temporarily enable `Sync deletion`.
2. Create a deletion scenario with one test contact.
3. Verify delete logic/prompts still work.
4. Disable `Sync deletion` again and re-check prompt suppression.

### Pass Criteria
- No deletion dialogs when `Sync deletion = OFF`.
- No repeated update loop for unchanged contacts.
- Match links remain stable across runs/restarts.
- Legitimate contact edits still sync correctly.
