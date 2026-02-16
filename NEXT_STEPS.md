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
