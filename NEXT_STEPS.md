## Next Steps

1. Install prerequisites on your build machine:
- Visual Studio 2022 (17.x) with `.NET desktop development`
- `.NET Framework 4.8.1` targeting/developer pack
- Outlook desktop (same bitness as target build, usually x86 for this app profile)

2. Build app-only private binary (no MSI/WiX required):
```powershell
cd C:\Utils\GO_Contact_Sync_Mod_fix\GOContactSyncMod_r1587
powershell -ExecutionPolicy Bypass -File .\scripts\build-private.ps1 -Configuration Release -Platform x86
```

3. Run your targeted regression checks:
- Use `TEST_PLAN.md` in this repo.

4. Apply this patch to another checkout/fork if needed:
```powershell
git apply .\patches\contact-resync-fix-r1587.patch
```

## What was prepared for you

- Patch artifact: `patches/contact-resync-fix-r1587.patch`
- Build script: `scripts/build-private.ps1`
- Test checklist: `TEST_PLAN.md`

## Tomorrow Priority: Startup "Loading Outlook folders" UX fix

Problem:
- App can appear hung on startup while scanning Outlook folders.
- Current behavior is confusing even when work is still progressing.

Goal:
- Never look stuck during folder discovery.
- Show clear progress and recovery guidance.

Proposed implementation:
1. Add a visible startup progress indicator with elapsed time + scanned folder count.
2. Add periodic heartbeat log/UI updates while scanning.
3. Add "Cancel scan" option and fail-safe timeout path.
4. Add a first-run prompt to choose scan mode:
- `Fast`: default store only
- `Full`: all stores/folders
5. Persist chosen scan mode per profile.

Acceptance criteria:
- On startup, user always sees active progress within 1 second.
- If scan exceeds threshold (for example 30s), UI shows actionable message (continue/cancel/switch to fast mode).
- No scenario where window is open but static with no status changes.
