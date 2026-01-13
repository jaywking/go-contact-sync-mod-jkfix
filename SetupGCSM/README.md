# +++ NEWS +++ NEWS +++ NEWS +++

# Version [4.2.1] - 14-04-2025
**r1585
FIX: added default folder https://sourceforge.net/p/googlesyncmod/bugs/1363/

# Version [4.2.0] - 01-03-2025
**r1584: 
updated to latest nugets

# Version [4.1.33] - 23-03-2023
**r1577 - 1583: 
updated to latest nugets
code refactoring (e.g. to remove the IsDeleted check for appointments)
Added configuration to SyncPastReminders

# Version [4.1.32] - 06-08-2022
**r1575 - 1576: 
FIXED: https://sourceforge.net/p/googlesyncmod/bugs/1350/ Contacts Skipped Due to No Unique Property Found and skipping Photo Sync
FIXED: https://sourceforge.net/p/googlesyncmod/bugs/1347/ Possible explanation for a number of the reported issues - Google Limit of 1000 contacts, app showing that contacts were deleted that were not, contacts being duplicated, performance being slow, etc.
ENHANCED: added checkbox to enable/disable contact photo sync, see following feature requests, and also to aovid the Google limit for Photo updates
https://sourceforge.net/p/googlesyncmod/feature-requests/206/
https://sourceforge.net/p/googlesyncmod/feature-requests/158/
https://sourceforge.net/p/googlesyncmod/feature-requests/97/
https://sourceforge.net/p/googlesyncmod/feature-requests/22/

# Version [4.1.31] - 29-07-2022
**r1566 - 1569: 
FIX: fixed Contacts Skipped Due to No Unique Property Found, see https://sourceforge.net/p/googlesyncmod/bugs/1350/
IMPROVEMENTS: Added support for the four missing types of phone number supported by Outlook.

## Version [4.1.30] - 21-07-2022
**r1564 - 1565: 
FIX: fixed object reference exception, see https://sourceforge.net/p/googlesyncmod/bugs/1352/
FIX: fixed startup delay from next sync + 5 minutes to only 5 minutes, see https://sourceforge.net/p/googlesyncmod/bugs/1325/


## Version [4.1.29] - 16-07-2022
**r1561 - 1563: 
FIX: made GCSM more case insensitive when comparing these labels, see https://sourceforge.net/p/googlesyncmod/bugs/1348/
IMPROVEMENTS: added RecRecreate-Functionality for deleted appointments, if the appointment to be deleted has multiple participants, don'T delete, but recreate on other side

## Version [4.1.28] - 28-01-2022
### SVN commits
**r1559 - 1560: 
FIX: Restructured the handling how and when to delete a contact, especially for SyncOptions GoogleToOutlookOnly and OutlookToGoogleOnly, might solve the following problems reported:
https://sourceforge.net/p/googlesyncmod/bugs/1324/
https://sourceforge.net/p/googlesyncmod/bugs/1299/
https://sourceforge.net/p/googlesyncmod/bugs/1296/
https://sourceforge.net/p/googlesyncmod/support-requests/883/


## Version [4.1.27] - 22-01-2022
### SVN commits
**r1555 - 1558:  
FIX: ignore exception when updating invalid recurrence interval for appointments
IMPROVEMENTS: Allow negative MonthsInFuture and also 0 MonthsInFuture or MonthsInPast to avoid resynching past reminders (see feature-requests #246 Should Past Events Still Trigger A Message)
IMPROVEMENTS: don't synchronize reminders of already finished events (see support-requests #879)
IMPROVEMENT: Added configuration "Only sync when idle"
IMPROVEMENT: Added the installed and available version number to the log string (to better trace, what was installed and what was the newest available from GCSM download)

## Version [4.1.25] - 20-11-2021
### SVN commits
**r1553 - 1554:  
FIX: freeze/lock of GUI thread (see bugs:#1324 and bugs:#1327 and 

## Version [4.1.23] - 05-11-2021
### SVN commits
**r1551 - 1552:  
- FIX: improved contact phone number handling, see bugs #1317

## Version [4.1.22] - 04-11-2021
### SVN commits
**r1547 - 1550:  
- IMPROVEMENT: Only save Contacts when changed (to avoid unnecessary logs, though nothing changed)
    - ToDo: Also apply above to appointments, but much more complicated, because too many Save methods and also RecurringEvents to be considered
- FIX: Fixed TestSyncGroups for Contacts
- IMPROVEMENT: remove seconds from OutlooProperties-LastSync, because cannot be retrieved again in most cases, therefore makes it more unaccurate to compare => always ignore seconds when comparing lastSync with lastUpdate
- FIX: index out of range
- FIX: Sync Thread abort if running
- FIX: GetRecurrence issues
- Updated to latest GoogleAPI nugets


## Version [4.1.20] - 21.10.2021
### SVN commits
**r1545 - 1546:  
- FIX: Handle empty WebPage type on Google side see [bug:##1319]
- FIX: Avoid HTTP Request exception, if there is a UserDefined attribute (e.g. language) with empty value
- IMPROVEMENT: Try to save Outlook items twice (sometimes the first time fails with COMException), use the same Save method for all Outlook items, and also check if Saved before Save

## Version [4.1.19] - 18.10.2021
### SVN commits
**r1543 - 1544:  
FIX: better user info message if missing folder selected in Settings GUI see [bug:#1314] and [support-requests:#878]

## Version [4.1.18] - 15.10.2021
### SVN commits
**r1540 - 1542:  
FIX: missing EndDate for recurrences
FIX: fixed the large body contact UnitTests
FIX: improved the handling of deleted recurrence exceptions

## Version [4.1.15] - 09.10.2021
### SVN commits
**r1536 - 1539:  
- FIX: [bugs:#1312] 
- removed seconds from AppointmentLastUpdated, because seconds are not saved in LastSync property anymore also for Appointments (same as for Contacts already in the past)
- added log entries when snycing recurrence exceptions
- FIX: fixed some conflict resolutions

## Version [4.1.13] - 07.09.2021
### SVN commits
**r1531 - 1535:  
- updated nuget for Google APIs to newest version  
- FIX: Added Polly based retry mechanism in LoadGoogleContacts [support request: #847]  
- FIX: Check for null in birthday fields as in some cases only Birthday. Text is set and Birthday.Date is null
- FIX: Fixed NullReference on Groups
- Added Info for using pre-release version. 
- FIX: Fixed some wrong "GO People Sync -> GO Contact Sync"

## Version [4.1.12] - 07.09.2021
### SVN commits
**r1519 - r1528:  
- updated nuget for Google APIs to newest version  
- FIX: Added retry mechanism for Contact writes [support request: #847]  

## Version [4.1.11] - 23.07.2021
**r1518:  
- fixed several places maybe causing the deleted contacts, sometimes there are magically the same resource twice in the contacts loaded from Google People API  
- updated nuget for Google APIs to newest version  

## Version [4.1.10] - 13.07.2021

### SVN commits
**r1502 - r1516:

- fixed RemoveGoogleDuplicatedContacts() to find "hanging/stops" problem 
- check for contact photo, don't return profile photo 
- fixed backward compliance with old id and new id for Google People API (to avoid deleting appointments)
- don't update global GoogleContacts for reloading a single contact
- fix get Contacts again
- fixed recurrence sync
- changed debug message move contact index to front of string
- fixed handling of recurrence exceptions
- fixed some handling about DuplicateGoogleAppointment
- improved logging, not show success log entry after exception updating appointment recurrence instance
- added version info into upload script template
- [SetupGCSM] changed $(SolutionDir) variable the use absolute path to compile from command line with msbuild
- [SetupGCSM] only build en-us installer culture
- [SetupGCSM] added pscp.exe (putty scp client) for uploading of releases
- [SetupGCSM] building a zip package after building the msi
- [SetupGCSM] added powershell script to uploading msi, zip and xml files to sourceforge file release system


## Version [4.0.9] - 21.06.2021

### SVN commits
**r1500 - r1501:
- added additional log entries to trace down [googlesyncmod:bugs] #1285 (e.g. to investigate root cause for https://sourceforge.net/p/googlesyncmod/support-requests/834/)
- added more debug info to UpdateRecurrence to investigate root cause for https://sourceforge.net/p/googlesyncmod/bugs/1280/


## Version [4.0.8] - 21.06.2021

### SVN commits
**r1496 - r1499:
- moved debug info from Info into Debug log (to avoid spamming the end user)
- added debug log to RemoveGoogleDuplicatedContacts (e.g. to investigate root cause for https://sourceforge.net/p/googlesyncmod/support-requests/834/)
- added more debug info to UpdateRecurrence to investigate root cause for https://sourceforge.net/p/googlesyncmod/bugs/1280/


## Version [4.0.7] - 19.06.2021

### SVN commits
**r1494 - r1495:
- fixed bug: https://sourceforge.net/p/googlesyncmod/bugs/1286/


## Version [4.0.6] - 16.06.2021

### SVN commits
**r1493:
- fixed: keep appointment private, if synchronized back with PrivateSetting=true
- updated to new Nuget packages
- fixed some minor bugs
- added additional logs for skipped entries (see https://sourceforge.net/p/googlesyncmod/support-requests/830/)

**r1492:
- changed update message label text
- comment out unused code for using own client secrets file
- HttpClient use only TLS 1.2 for update checks
- creating release zip file (portable version of gcsm) for installer project


## Version [4.0.5] - 12.06.2021

### SVN commits

**r1491:
- fixed bug #825 object reference error
- replaced several unnecessary usages of nullable ContactPropertiesUtils.Get...


## Version [4.0.4] - 09.06.2021

### SVN commits

**r1490:
- fixed bugs #1278 object reference error
- added Utc2Local conversion for all dates

## Version [4.0.3] - 08.06.2021

### SVN commits

**r1484 - r1489:
- Fixed several bugs
- Fixed the UnitTests and got them running again

## Version [4.0.1] - 07.06.2021

### SVN commits

**r1462 - r1482:
- Fixed several bugs
- Fixed the UnitTests and got them running again
- Migrated GoogleContact GData API to new Google People API
- New versions of NuGet packages.

## Version [3.11.17] - 23.03.2021

### SVN commits

**r1455 - r1461:

- FIX: Handle appointments with empty updated date [bug: #1273]
- New versions of NuGet packages.

## Version [3.11.16] - 26.01.2021

### SVN commits

**r1448 - r1454:

- FIX: Properly initialize day of month [bug: #1266]
- New versions of NuGet packages.

## Version [3.11.15] - 21.01.2021

### SVN commits

**r1445 - r1447:

- FIX: Properly initialize day of week mask [bug: #1266]

## Version [3.11.14] - 20.01.2021

### SVN commits

**r1442 - r1444:

- FIX: Debugging for incorrect day of week mask [bug: #1266]

## Version [3.11.13] - 17.01.2021

### SVN commits

**r1432 - r1441:

- FIX: Fixed detection of Outlook appointments outside sync range [bug: #1261]
- FIX: Handle "Attempted to read or write protected memory" errors coming from very old versions of Outlook [support request: #802]
- FIX: Use Polly to handle timeouts [support request: #800]
- New versions of NuGet packages.

## Version [3.11.12] - 28.12.2020

### SVN commits

**r1429 - r1431:

- FIX: Fixed detection of Outlook appointments outside sync range [bug: #1261]

## Version [3.11.11] - 27.12.2020

### SVN commits

**r1426 - r1428:

- FIX: Fixed detection of Outlook appointments outside sync range [bug: #1261]

## Version [3.11.10] - 27.12.2020

### SVN commits

**r1418 - r1425:

- FIX: Fixed detection of Outlook appointments outside sync range [bug: #1261]
- FEATURE: Removed job title from contacts summary [feature: #215]
- New versions of NuGet packages.

## Version [3.11.9] - 01.12.2020

### SVN commits

**r1415 - r1417:

- FIX: Better conversion from string to DateTime? [bug: #1260]

## Version [3.11.8] - 01.12.2020

### SVN commits

**r1412 - r1414:

- FIX: Better conversion from string to DateTime? [bug: #1259]

## Version [3.11.7] - 29.11.2020

### SVN commits

**r1409 - r1411:

- FIX: Better conversion from string to DateTime? [bug: #1259]

## Version [3.11.6] - 29.11.2020

### SVN commits

**r1406 - r1408:

- FIX: Added logging for incorrect PatternEndDate [bug: #1258]

## Version [3.11.5] - 27.11.2020

### SVN commits

**r1402 - r1405:

- FIX: Added logging for incorrect PatternEndDate [bug: #1258]

## Version [3.11.4] - 19.11.2020

### SVN commits

**r1398 - r1401:

- FIX: Do not change focus in case of 0x800401E3 [bug: #1256]
- New versions of NuGet packages.

## Version [3.11.3] - 15.11.2020

### SVN commits

**r1384 - r1397:

- FEATURE: Block setting auto sync more frequently then 1 hours (switch field in GUI from minutes to hours)
- FEATURE: Allow users to provide their own client secrets
- FIX: Better handling of exceptions while handling Outlook appointments throwing exceptions (for example due to incorrect birthday date set) [support request: #787]
- FIX: Modify check in installer script to validate minimal OS version (Windows 7) and .NET version (4.6.1)
- FIX: More clear error message in case of appointment folder is not selected [bug: #1253]
- New versions of NuGet packages.

## Version [3.11.2] - 26.10.2020

### SVN commits

**r1374 - r1383:

- FIX: Improved looging of appointment exceptions [support request: #784]
- FIX: Improved logging for errors with Contacts
- FIX: Remove spurious spaces in GetTitleFirstLastAndSuffix
- FIX: Better logging of contacts with empty names [support request: #783]
- New versions of NuGet packages.

## Version [3.11.1] - 19.10.2020

### SVN commits

**r1369 - r1373:

- FIX: Resync google appointments before updating them [support request: #782]
- FIX: Increased rolling log file size to 10M bytes

## Version [3.11.0] - 18.10.2020

### SVN commits

**r1365 - r1368:

- FIX: Better logging of strings with huge contents [support request: #778]
- FEATURE: Switched logging to Serilog framework

## Version [3.10.108] - 08.10.2020

### SVN commits

**r1362 - r1364:

- FIX: Better logging of strings with huge contents [support request: #778]
- New versions of NuGet packages.

## Version [3.10.107] - 08.10.2020

### SVN commits

**r1357 - r1361:

- FIX: Better logging of strings with huge contents [support request: #778]
- FIX: Better handling of 800401E3 COMException [bug: #1243]
- FIX: Better logging of 403 errors [bug: #1248]

## Version [3.10.106] - 06.10.2020

### SVN commits

**r1347 - r1356:

- FEATURE: Remember windows positions and size
- FIX: Better handling of 80040201 COMException [bug: #1243]
- FIX: Better logging of errors in case user is not organizer of the appointment [support request: #772]
- FIX: Better logging of contact info in case contact has no names but only company name [support request: #772]
- New versions of NuGet packages.

## Version [3.10.105] - 29.09.2020

### SVN commits

**r1343 - r1346:

- FIX: Skipping contacts returning Invalid country code: ZZ [support request: #772]
- New versions of NuGet packages.

## Version [3.10.104] - 29.09.2020

### SVN commits

**r1337 - r1342:

- FIX: Skipping contacts returning Invalid country code: ZZ [support request: #772]
- FIX: Better handling of 80040201 COMException [bug: #1243]
- FIX: Handling exceptions while looking for items out of sync range [bug: #1247]
- FIX: Handle exceptions in ToLogString [support request: #774]

## Version [3.10.103] - 28.09.2020

### SVN commits

**r1332 - r1336:

- FIX: Resize dialog window properly [support request: #770]
- New versions of NuGet packages.

## Version [3.10.102] - 23.09.2020

### SVN commits

**r1328 - r1331:

- FIX: Better handling of 80040201 COMException [bug: #1243]
- New versions of NuGet packages.

## Version [3.10.101] - 22.09.2020

### SVN commits

**r1325 - r1327:

- FIX: Better matching for full day appointments [bug: #1222]

## Version [3.10.100] - 21.09.2020

### SVN commits

**r1320 - r1324:

- FIX: Handle "An object cannot be found." COMException when looking for appointments out of sync range [bug: #1241]
- FIX: For some appointments setting StartUTC throws COMException "The object does not support this method" [bug: #1241]
- New versions of NuGet packages.

## Version [3.10.99] - 20.09.2020

### SVN commits

**r1299 - r1317:

- FIX: Handle exceptions while trying to access AppointmentItem inside Outlook appointment exceptions [support request: #761]
- FIX: Cosmetic change of label [feature request: #242]
- FIX: Retry in case during saving photo we got 404 error [bug: #1238]
- FIX: Load appointments out of sync range in case not found during match process [bug: #1222]
- FIX: Relogon to Outlook in case of 80040201 COMException [bug: #975]
- FIX: Skip synchronization of appointments not possible to create in Outlook [support request: #756]
- FIX: Remove check for size of notes field, now app tries to save and in case error from Google API displays warning [support request: #755]
- FIX: Fixed saving of Force RTF Appointments [bug: #1233]
- New versions of NuGet packages.

## Version [3.10.98] - 20.07.2020

### SVN commits

**r1294 - r1298:

- FIX: Correct info about installed Office version [support request: #753]
- FIX: Layout fixes [bug: #1229]
- New versions of NuGet packages.

## Version [3.10.97] - 10.07.2020

### SVN commits

**r1290 - r1293:

- FIX: Ignore RequestError errors in case of undeleting events [support request: #752]
- New versions of NuGet packages.

## Version [3.10.96] - 09.07.2020

### SVN commits

**r1285 - r1289:

- FIX: Some contacts despite having less extendedProperties still can throw such exception [support request: #750]
- FIX: Add more logging in case of RequestError Forbidden [403] [support request: #752]
- FIX: Move old log files to "Archived Logs" directory and have different name of archived log

## Version [3.10.95] - 08.07.2020

### SVN commits

**r1282 - r1284:

- FIX: Fixed sync between non recurring Outlook and recurring Google [bug: #1226]

## Version [3.10.94] - 08.07.2020

### SVN commits

**r1279 - r1281:

- FIX: Add more debugging [bug: #1226]

## Version [3.10.93] - 08.07.2020

### SVN commits

**r1268 - r1278:

- FIX: Add more debugging [bug: #1226]
- FIX: Make all dialogs TopMost windows  [bug: #1215]
- FIX: Deleting currently used profile will revoke Google auth token [bug: #1186]
- FIX: Layout changes to ConfigurationManager window [bug: #1186]
- New versions of NuGet packages.

## Version [3.10.92] - 29.06.2020

### SVN commits

**r1263 - r1267:

- FIX: Added more logging for errors triggered in setting Sensitivity [bug: #1223]
- FIX: Ignore shorter User Properties [bug: #1206]

## Version [3.10.91] - 26.06.2020

### SVN commits

**r1260 - r1262:

- FIX: Added null check in IsSameRecurrenc [bug: #1222]

## Version [3.10.90] - 26.06.2020

### SVN commits

**r1257 - r1259:

- FIX: Retrieve Google appointment before update in case of updates to appointment with exception [bug: #1221]

## Version [3.10.89] - 25.06.2020

### SVN commits

**r1253 - r1256:

- FIX: Avoid accessing null dates in case of updates to appointment with exception [bug: #1220]

## Version [3.10.88] - 25.06.2020

### SVN commits

**r1249 - r1252:

- FIX: Avoid accessing null dates in case of updates to appointment with exception [bug: #1220]
- New versions of NuGet packages.

## Version [3.10.87] - 23.06.2020

### SVN commits

**r1243 - r1248:

- FIX: Better handling of recurrence in Office 365 [support: #746]
- FIX: Improved logging  [bug: #1216]
- FIX: Avoid unnecessary logging for multiple participants appointments [support: #719]
- New versions of NuGet packages.

## Version [3.10.86] - 17.06.2020

### SVN commits

**r1240 - r1242:

- FIX: Ignore exceptions when trying to access recurrence exception appointment item [bug: #1214]

## Version [3.10.85] - 16.06.2020

### SVN commits

**r1237 - r1239:

- FIX: Fixed changes to recurrence state  [bug: #1212]

## Version [3.10.84] - 15.06.2020

### SVN commits

**r1231 - r1236:

- FIX: Improved logging  [bug: #1212]
- New versions of NuGet packages.

## Version [3.10.83] - 10.06.2020

### SVN commits

**r1203 - r1230:

- FIX: Accessing RTFBody sometimes throw AccessViolationException [bug: #1207]
- FIX: Add more logging in case synchronizing empty contacts [bug: #1210]
- FIX: Accessing StartInStartTimeZone and EndInEndTimeZone in older version of Outlook (like 2003) throw AccessViolationException  [bug: #1209]
- FIX: Better handling of appointment recurrence deleted exceptions [bug: #1184]
- FIX: Treat different forms of Google email as same [bug: #1184]
- New versions of NuGet packages.

## Version [3.10.82] - 03.04.2020

### SVN commits

**r1195 - r1202:

- FIX: Handle "Not a legal OleAut date" exception  [bug: #1197]
- New versions of NuGet packages.

## Version [3.10.81] - 22.03.2020

### SVN commits

**r1190 - r1194:

- FIX: Google Appointment with OriginalEvent found, but Outlook occurrence not found [bug: #1184]

## Version [3.10.80] - 21.03.2020

### SVN commits

**r1188 - r1189:

- FIX: Google Appointment with OriginalEvent found, but Outlook occurrence not found [bug: #1184]

## Version [3.10.79] - 18.03.2020

### SVN commits

**r1182 - r1187:

- FIX: Google Appointment with OriginalEvent found, but Outlook occurrence not found [bug: #1184]
- FIX: Treat different forms of Google email as same [bug: #1184]
- New versions of NuGet packages.

## Version [3.10.78] - 10.03.2020

### SVN commits

**r1174 - r1181:

- FIX: Treat different forms of Google email as same [bug: #1184]
- New versions of NuGet packages.

## Version [3.10.77] - 02.03.2020

### SVN commits

**r1162 - r1173:

- FIX: Accessing user properties by using [] operator is case sensitive,
  but later Add fails to add new property as it is case insensitive [bug: #1184]

## Version [3.10.76] - 05.02.2020

### SVN commits

**r1158 - r1161:

- FIX: Accessing user properties by using [] operator is case sensitive,
  but later Add fails to add new property as it is case insensitive [bug: #1184]
  
## Version [3.10.75] - 04.02.2020

### SVN commits

**r1154 - r1157:

- FIX: Replicate the fix from 3.10.74, but for appointments [bug: #1184]

## Version [3.10.74] - 03.02.2020

### SVN commits

**r1150 - r1153:

- FIX: Accessing user properties by using [] operator is case sensitive,
  but later Add fails to add new property as it is case insensitive [bug: #1184]
- New versions of NuGet packages.

## Version [3.10.73] - 01.02.2020

### SVN commits

**r1147 - r1149:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.72] - 01.02.2020

### SVN commits

**r1143 - r1146:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.71] - 30.01.2020

### SVN commits

**r1139 - r1142:

- FIX: Added debug logging to track why contacts are being updated
- New versions of NuGet packages.

## Version [3.10.70] - 30.01.2020

### SVN commits

**r1136 - r1138:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.69] - 30.01.2020

### SVN commits

**r1133 - r1135:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.68] - 29.01.2020

### SVN commits

**r1129 - r1132:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.67] - 29.01.2020

### SVN commits

**r1127 - r1128:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.66] - 28.01.2020

### SVN commits

**r1124 - r1126:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.65] - 27.01.2020

### SVN commits

**r1120 - r1122:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.64] - 26.01.2020

### SVN commits

**r1117 - r1119:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.63] - 26.01.2020

### SVN commits

**r1116 - r1116:

- FIX: Added debug logging to track why contacts are being updated

## Version [3.10.62] - 25.01.2020

### SVN commits

**r1107 - r1115:

- FIX: Added debug logging to track why contacts are being updated
- FIX: Load folders from settings in case of profile changes  [support: #707]
- FIX: Report some Outlook exceptions as hard errors
- FIX: Fixed synchronization of other fax number [bug: #1187]
- New versions of NuGet packages.
- FIX: Spelling mistake in log message

## Version [3.10.61] - 09.12.2019

### SVN commits

**r1103 - r1106:

- FIX: Normalized contact names when printing to log file. Added counter. [support: #703]

## Version [3.10.60] - 02.12.2019

### SVN commits

**r1098 - r1102:

- FIX: Add delay in case of retrieving contact one by one. Added more clear diagnostic. [bug: #1173]

## Version [3.10.59] - 01.12.2019

### SVN commits

**r1096 - r1097:

- FIX: Add delay in case of retrieving contact one by one. Added more clear diagnostic. [bug: #1173]

## Version [3.10.58] - 28.11.2019

### SVN commits

**r1093 - r1095:

- FIX: Add delay in case of retrieving contact one by one. Added more clear diagnostic. [bug: #1173]

## Version [3.10.57] - 27.11.2019

### SVN commits

**r1089 - r1092:

- FIX: Add delay in case of retrieving contact one by one. Added more clear diagnostic. [bug: #1173]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.56] - 17.11.2019

### SVN commits

**r1083 - r1088:

- FIX: Use Google People API to get more details in case of parse exception errors [bug: #1165]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.55] - 31.10.2019

### SVN commits

**r1078 - r1082:

- FIX: Don't block the main thread if the user aborts (or timeout) the google authorization.
- FIX: Save current settings when user selects different profile [bug: #1160]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.54] - 19.10.2019

### SVN commits

**r1071 - r1077:

- FIX: Use Google People API to get more details in case of parse exception errors [bug: #1165]
- FIX: Do not call count method to avoid 0x0x800706BA error [bug: #1166]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.53] - 08.09.2019

### SVN commits

**r1063 - r1070:

- FIX: Fix resizing problem with warning forms [bug: #1037]
- FIX: Retry in case 0x80010001 COM error received [bug: #1158]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.52] - 11.08.2019

### SVN commits

**r1059 - r1062:

- FIX: Change conversion algorithm of RTF notes in contacts and appointments [bug: #1157]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.51] - 28.07.2019

### SVN commits

**r1051 - r1058:

- FIX: Added retry in case of network connectivity issues [bug: #1016]
- FIX: Change level of warnings for one way sync [bug: #1149]
- FIX: Release memory for TimeZone COM objects
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.50] - 14.06.2019

### SVN commits

**r1047 - r1050:

- FIX: For OutlookToGoogleOnly always create Google contact if does not exist [bug: #1101]
- FIX: Improve memory usage

## Version [3.10.49] - 13.06.2019

### SVN commits

**r1039 - r1046:

- FIX: Add retries in case of ThreadAbortException while loading Google contacts [support: #666]
- FIX: Improved logging of Google exceptions
- FIX: Limit logging in case no matches
- FIX: Improved logging of google appointment exceptions
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.48] - 06.06.2019

### SVN commits

**r1036 - r1038:

- FIX: Check if Reminders are not null [bug: #1146]

## Version [3.10.47] - 05.06.2019

### SVN commits

**r1031 - r1035:

- FIX: Improved logging of COM exceptions [bug: #1143]
- FIX: Rename enum to Outlook_2016_or_2019_or_365
- FIX: Do not check for organizer of events, so typical annual events will also be synced [bug: #1129]

## Version [3.10.46] - 03.06.2019

### SVN commits

**r1026 - r1030:

- FIX: Check if appointment is recurrent when accessing recurrence info [bug: #1134]
- FIX: Check if appointment is recurrent when logging info [bug: #1134]
- FIX: Improved logging of Google exceptions [bug: #1129]

## Version [3.10.45] - 01.06.2019

### SVN commits

**r1022 - r1025:

- FIX: In case old version of Outlook update Start/End instead of StartUTC/EndUTC  [bug: #1129]
- FIX: Improved logging of Google exceptions [bug: #1129]

## Version [3.10.44] - 31.05.2019

### SVN commits

**r1018 - r1021:

- FIX: Log exception details [bug: #1129]
- FIX: Removed logging of all appointments [bug: #1134]

## Version [3.10.43] - 30.05.2019

### SVN commits

**r1011 - r1017:

- FIX: Ignore Outlook appointments with all exceptions deleted [bug: #1127]
- FIX: Do not access Conflicts property [bug: #1134]
- FIX: Try to re-login in case of 80040201 exception [bug: #1135]
- FIX: EndUTC is not available in old versions of Outlook (like 2003) [bug: #1129]

## Version [3.10.42] - 29.05.2019

### SVN commits

**r1006 - r1010:

- FIX: Accessing ConversationIndex and ConversationTopic for some appointments throws Member not found exception (0x80020003) [bug: #1134]
- FIX: StartUTC is not available in old versions of Outlook (like 2003) [bug: #1129]
- FIX: Ignore exceptions when logging [bug: #1127]

## Version [3.10.41] - 28.05.2019

### SVN commits

**r1000 - r1005:

- FIX: In logfile change Outlook2016 to Outlook2016_or_2019_or_365 [bug: #1135]
- FIX: For 80040201 exception try to re-login to Outlook [bug: #1135]
- FIX: Added logging for exception thrown while updating some appointments in old Outlook (2003) [bug: #1129]
- FIX: Accessing ConversationID for some appointments throws Member not found exception (0x80020003) [bug: #1136]

## Version [3.10.40] - 27.05.2019

### SVN commits

**r996 - r999:

- FIX: Accessing ConversationID for some appointments throws Member not found exception (0x80020003) [bug: #1134]
- FIX: Improved detection of Office365 [bug: #1131]

## Version [3.10.39] - 26.05.2019

### SVN commits

**r984 - r995:

- FIX: Added logging for CheckVersion calls [bug: #1133]
- FIX: Added logging for empty Google contacts [bug: #1132]
- FIX: Added logging for Outlook exceptions [bug: #1127]
- FIX: Better logging of skipped Outlook contacts [bug: #1130]
- FIX: StartTimeZone and EndTimeZone are not available in old versions of Outlook (like 2003) [bug: #1129]

## Version [3.10.38] - 20.05.2019

### SVN commits

**r976 - r983:

- FIX: Better logging of empty google appointment exceptions [bug: #1127]
- FIX: Handle exceptions when accessing Google appointment exception
- FIX: Avoid AccessViolationException for old Office versions (2003) [bug: #1129]
- FIX: Force TLS 1.2 for some rare Windows setups [bug: #1125]
- FIX: Improved detection of Click-To-Run Office installations [bug: #1124]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.37] - 23.04.2019

### SVN commits

**r968 - r975:

- FIX: Namespace.Accounts available only in newer version of Outlook [bug: #1121]
- FIX: Make error reporting more robust against potential exceptions [support: #657]
- FIX: Additional check for registry [support: #656]
- FIX: Change limit for number of contact extended properties [bug: #1108]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.36] - 12.04.2019

### SVN commits

**r965 - r967:

- FIX: Additional check for registry [support: #656]

## Version [3.10.35] - 11.04.2019

### SVN commits

**r961 - r964:

- FIX: Minimize memory consumption during processing [bug: #1119]
- UPDATE: Update NuGet packages to latest version.

## Version [3.10.34] - 09.04.2019

### SVN commits

**r958 - r960:

- FIX: Fixed NullReferenceException [bug: #1116]

## Version [3.10.33] - 07.04.2019

### SVN commits

**r947 - r957:

- FIX: Optimized memory usage [bug: #1107]
- FIX: Load google appointments combo in case such folder is stored in registry [bug: #1106]
- FIX: Improved handling of COMErrors when creating Outlook.Application object [bug: #1081]

## Version [3.10.32] - 08.02.2019

### SVN commits

**r943 - r946:

- FIX: Revert restrict mode for items retrieved from Outlook [bug: #1099]

## Version [3.10.31] - 06.02.2019

### SVN commits

**r939 - r942:

- FIX: Revert restrict mode for items retrieved from Outlook [bug: #1085]
- FIX: Log error code for SEHException [bug: #1097]

## Version [3.10.30] - 05.02.2019

### SVN commits

**r935 - r938:

- FIX: Handle 0x800706BE exception [bug: #1093]
- FIX: Fix Restrict sorting if condition is based only on End date [bug: #1091]

## Version [3.10.29] - 05.02.2019

### SVN commits

**r929 - r934:

- FIX: Handle loop in case exceptions from appointments [bug: #1089]
- FIX: Handle null appointments [bug: #1089]
- FIX: Handle COM Error 80040109 [bug: #1090]

## Version [3.10.28] - 04.02.2019

### SVN commits

**r925 - r928:

- FIX: Handle COM Error 80040109 [bug: #1088]

## Version [3.10.27] - 03.02.2019

### SVN commits

**r919 - r924:

- FIX: Restrict number of items retrieved from Outlook [bug: #1085]
- FIX: Avoid deleting temporary photo file if it is used [bug: #1086]
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.26] - 21.01.2019

### SVN commits

**r915 - r918:

- FIX: Reverted back to 3.10.22

## Version [3.10.25] - 20.01.2019

### SVN commits

**r912 - r914:

- FIX: Restrict number of items retrieved from Outlook [bug: #1078]

## Version [3.10.24] - 19.01.2019

### SVN commits

**r907 - r911:

- FIX: Restrict number of items retrieved from Outlook [bug: #1078]
- FIX: Fix removing items from collection [bug: #901]
- FIX: Keep match even if one of the item (Outlook appointment or Google event) is outside sync range [bug: #1066]

## Version [3.10.23] - 14.01.2019

### SVN commits

**r904 - r906:

- FIX: Restrict number of items retrieved from Outlook [bug: #1078]

## Version [3.10.22] - 12.01.2019

### SVN commits

**r885 - r903:

- FIX: Improve logging for 412 Google errors [bug: #1076]
- FIX: Added support for more types of address entries on Exchange
- FIX: Improved log reporting for different event organizers at Google
- FIX: Handle incorrect date values in LastSync field [bug: #1074]
- FIX: Improved log reporting for Google contacts with empty title
- FIX: Reload Google appointment folders combo if user name was changed [bug: #1072]
- FIX: Removed not necessary ReleaseComObject calls
- FIX: Make "Empty Outlook contact found" as Debug message not Warning
- FIX: Added detection of Office 2019 or Office 365
- FIX: Added support for more types of address entries on Exchange
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.21] - 17.12.2018

### SVN commits

**r879 - r884:

- FIX: Skip logging of Outlook contact details if it triggers exception [bug: #1067]

## Version [3.10.20] - 17.12.2018

### SVN commits

**r868 - r879:

- FIX: Removed workaround [bug: #1062]
- FIX: Better handling of new versions of Outlook [bug: #1067]
- FIX: Additional logging when ArgumentException [bug: #1063]
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.19] - 01.11.2018

### SVN commits

**r849 - r867:

- FIX: Photos synchronizaton between Outlook and Google [bug: #1062]
- FIX: Do not change Display Name for emails at Outlook [bug: #1034]
- FIX: Warn if Google contact has custom phone label [bug: #1042]
- FIX: Avoid exception while accessing Outlook contact data during previous exception logging [bug: #1045]
- FIX: Handle ProtocolViolationException while saving Google Groups [bug: #1005, #1036]
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.18] - 18.10.2018

### SVN commits

**r831 - r848:

- FIX: Improved debugging of errors with connecting to Outlook
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.17] - 20.09.2018

### SVN commits

**r825 - r827:

- FIX: Set TLS to 1.2 in order to correctly download XML file from SF [bug: #1014]
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.16] - 16.08.2018

### branched from revision r807 to 3.10.16 branch and merged newest ContactSync

- FIX: SetupGCSM-3.10.15 NOT syncing MS Outlook Calendar " test " appointment to Google [bug: #1028]
- Reverted Refactoring from Obelix for above FIX

## Version [3.10.15] - 14.08.2018

### SVN commits

**r801 - r819:

- FIX: Handle "Object reference not set to an instance of an object" if Country is null on Google side [bug: #1027]
- FIX: Use dates instead number of months for appointment synchronization range
- FIX: Use random temporary file name for contact photos
- FIX: Use consistent time range for synchronization [bug: #940]
- FIX: More precise logging message
- Refactoring for potential split of synchronization options between contacts and appointments
- UPDATE: Update NuGet packages to latest version, incl. new version of Google API

## Version [3.10.14] - 27.11.2016

### SVN commits

**r790 - r799:

- FIX: Do not scan for duplicates during matching, duplicates were removed during load [bug: #954]
- FIX: Fixed NullReferenceException in Office version checking
- FIX: Handle DPI resolutions [bug: #953]
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.13] - 20.11.2016

### SVN commits

**r787 - r789:

- FIX: Avoid further matching in case appointments are already matched [bug: #951]

## Version [3.10.12] - 19.11.2016

### SVN commits

**r781 - r786:

- FIX: Check if value of extended property field is null [bug: #950]

## Version [3.10.11] - 18.11.2016 

### SVN commits

**r765 - r780:

- FIX: Check if UserDefinedFields field is null [bug: #948]
- FIX: Check if FormattedAddress field is null [bug: #945]
- FIX: Workaround for issue in Google client libraries [bug: #866]
- IMPROVEMENT: Do not warn about skipping empty contact if this is distribution list
- FIX: Implement OleMessageFilter to handle RPC_E_CALL_REJECTED errors [bug #939]
- FIX: Unhide label with status text [bug #942]
- FIX: Retry in case ProtocolViolationException exception during Google contact save [bug #903]
- FIX: Do not add custom field to folder [bug #651]

## Version [3.10.10] - 12.11.2016

### SVN commits

**r744 - r764:

- FIX: Preserve FileAs format when updating existing Outlook contact [bug #543]
- FIX: Change mapping between Outlook and Gmail for email types (do not use display name from Outlook) [bug #932]
- FIX: Added more logging [bugs #843, #897]
- FIX: Handle contacts with duplicated extended properties [bug #655]
- FIX: Handle contacts with too big extended properties [bug #895]
- FIX: Handle contacts with more than 10 extended properties [bug #900]
- FIX: Do not synchronize from Outlook phone numbers with only white spaces [bug #629]
- FIX: Add dummy values to contact user properties or contact extended properties [bugs #634, #886]
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.9] - 21.10.2016 

### SVN commits

**r737 - r743:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Fixed regression in selecting folders.

## Version [3.10.8] - 18.10.2016 

### SVN commits

**r727 - r736:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Add more logging for exceptions during accessing user properties [bug #651]
- FIX: Additional check to avoid access violation [bug #567]
- FIX: Fixed regression from 3.10.7, logon to MAPI using selected folder not default one.

## Version [3.10.7] - 15.10.2016 

### SVN commits

**r711 - r726:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Some users have emty time zone at Google, in such situation try to use what is set in GUI [bug #878]
- FIX: Do not throw exception in case there is problem with registry
- FIX: Layout changes for high DPI setups
- FIX: Handle situation when Outlook folder is invalid, logon to Outlook using default folders
- FIX: Handle situation when previously selected Outlook folder became invalid (for example was deleted in Outlook)

## Version [3.10.6] - 07.10.2016 

### SVN commits

**r705 - r710:

- FIX: Logon to MAPI in case of exception [bug #871]
- FIX: Select the first item in folder combo [bug #871]

## Version [3.10.5] - 06.10.2016 

### SVN commits

**r695 - r704:

- FIX: Add more logging [bugs #871, #877, #878, #879]
- FIX: Improved logging for COMExceptions [bug #871]
- FIX: In case folder was set not correctly, switch to default one [bug #871]
- FIX: Avoid exception if version information cannot be read 

## Version [3.10.4] - 04.10.2016 

### SVN commits

**r689 - r694:

- FIX: Release memory while scanning Outlook items [bug #874]
- FIX: Added more detailed logging [bug #871]
- FIX: Handle situation when bitness is not set in registry [bug #876]

## Version [3.10.3] - 03.10.2016 

### SVN commits

**r672 - r688:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Corrected time zone mapping between Google and Outlook
- FIX: Added more logging [bugs: #863, #870]
- FIX: Added AutoGenerateBindingRedirects

## Version [3.10.2] - 29.09.2016 

### SVN commits

**r660 - r671:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Set time zone for recurrent appointments.
- FIX: Added more logging [bugs: #870, #871]

## Version [3.10.1] - 25.09.2016 

### SVN commits

**r610 - r659:

- UPDATE: Update NuGet packages to latest version 
- FIX: update detection routine to fetch information about the latest version [bugs: #795, #826, #845, #853]
- FIX: Ignore exceptions when retrieving windows version, put more diagnostic to log in case of exception [bug: #849]
- FIX: Remove duplicates from Outlook: two different Outlook appointments pointing to the same Google appointment. [bug: #614]
- FIX: Force Outlook to set country in formatted address string. [bug: #850]
- FIX: Clear Google reminders in case Outlook appointment has no reminders. [bug: #599]
- FIX: Synchronize time zones. [bugs: #533, #654, #813, #851, #852, #856]

## Version [3.10.0] - 14.06.2016

### SVN commits

**r586 - r609**:

- CHANGE: Retargetted to .NET 4.5,  as a result Windows XP is not supported anymore, minimum requirement is Windows Vista SP2.
- UPDATE: Update NuGet packages to latest version (new version of Google client libraries require .NET 4.5)
- FIX: ResetMatches rewritten to use BatchRequest functionality [bugs: #673, #738, #796, #799, #806, #836]
- FIX: Warning in exception handler to indicate appointment which triggered error (feature request: #148)
- FIX: Log instead of Error Handler to avoid multiple Windows
- FIX: Added more info about raised exceptions

## Version [3.9.15] - 11.03.2016

### SVN commits

**r582 - r583**:

- FIX: merged back AppointmentSync to use ForceRTF
- UPDATE: Removed GoogleDocuments 2nd level authentication, because no notes sync possible currently (no need to provide GCSM access to GoogleDocuments)

## Version [3.9.14] - 24.02.2016

### SVN commits

**r572 - r579**:

- FIX: handle busy/free/tentative status by transparency, see <https://sourceforge.net/p/googlesyncmod/bugs/463/>
- FIX: implemented ForceRTF checkbox
- UPDATE: update NuGet packages to latest version
- UPDATE: tooltip (UserName)
- FIX: Added more diagnostics for problems with Outlook installation [bugs:#785]
- FIX: changed copy to clipboard code to prevent HRESULT: CLIPBRD_E_CANT_OPEN, see [bugs:#749]
- UPDATE: field label User -> E-Mail in UI FIX: changed error typo

## Version [3.9.13] - 01.11.2015

### SVN commits

**r567 - r571**:

- FIX: [bugs:#780]
- UPDATE: nuget packages
- IMPROVEMENT: detect version: Outlook 2016
- IMPROVEMENT: log windows version: name, architecture, number
- IMPROVEMENT: do not copy the interop dll to output dir
- IMPROVEMENT: do not include interop into setup
- CHANGE: target type to AnyCPU
- CHANGE: remove Office.dll (not necessary)
- IMPROVEMENT: Added notes how to repair VS2013 installation after modifying machine.config for UnitTests
- prepared new setup

## Version [3.9.12] - 16.10.2015

### SVN commits

**r563 - r566**:

- Reverted change from 3.9.11: Referenced Outlook 2013 Interop API and copied it locally
turned out, also not runnable with Outlook 2016
and has issues with Older Office 2010 and 2007 installations

## Version [3.9.11] - 15.10.2015

### SVN commits

**r558 - r562**:

- FIX: Workaround, to not overwrite tentative/free Calendar items, see [bugs:#709]
- FIX: [bugs:#731]
- UPDATE: nuget packages
- FIX: don't load old registry settings to avoid profile errors
- FIX: Remove recurrence from slave, if removed from master
- FIX: Extended ListSeparator for GoogleGroups
- FIX: handle exception when saving Outlook appointment fails (log warning instead of stop and throw error)

## Version [3.9.10] - 16.05.2015

### SVN commits

**r555 - r557**:

- FIX: Remove recurrence from slave, if removed from master
- FIX: Extended ListSeparator for GoogleGroups
- FIX: handle exception when saving Outlook appointment fails (log warning instead of stop and throw error)

## Version [3.9.9] - 12.05.2015

### SVN commits

**r552 - r553**

- FIX: Improved GUI behavior, if CheckVersion fails (e.g. because of missing internet connection or wrong proxy settings)
- FIX: added America/Phoenix to the timezone Dropdown

## Version [3.9.8] - 04.05.2015

### SVN commits

**r546 - r550**

- FIX: stopped duplicating Group combinations and adding them to Google, [see](https://sourceforge.net/p/googlesyncmod/bugs/691/)
- FIX: avoid "Forbidden" error message, if calender item cannot be changed by Google account, [see](https://sourceforge.net/p/googlesyncmod/bugs/696/)
- FIX: removed debug update detection code
- UPDATE: Google.Apis.Calendar.v3
- FIX: moving "Copy to Clipboard" back to own STA-Thread
- FIX: ballon tooltip for update was always shown (svn commit error)

## Version [3.9.7] - 21.04.2015

### SVN commits

**r542 - r544**

- FIX: Removed Notes Sync, because not supported by Google anymore
- FIX: Handle null values in Registry Profiles, [see](http://sourceforge.net/p/googlesyncmod/bugs/675/)

**Free Open Source Software, Hell Yeah!**

## Version [3.9.6] - 15.04.2015

### SVN commits

**r536 - r541**

- **IMPROVEMENT**: adjusted error text color
- **IMPROVEMENT**: Made Timezone selection a dropdown combobox to enable users to add their own timezone, if needed (e.g. America/Arizona)
- **IMPROVEMENT**: check for latest downloadable version at sf.net
- **IMPROVEMENT**: check for update on start
- **IMPROVEMENT**: added new error dialog for user with clickable links
- **FIX**: renamed Folder OutlookAPI to MicrosoftAPI
- **FIX**: <https://sourceforge.net/p/googlesyncmod/bugs/700/>
- **CHANGE**: small fixes and changes to the Error Dialog

**Free Open Source Software, Hell Yeah!**

## Version [3.9.5] - 10.04.2015

### SVN commits

**r535**

- **FIX**: Fix errors when reading registry into checkbox or number textbox, see https://sourceforge.net/p/googlesyncmod/bugs/667/
https://sourceforge.net/p/googlesyncmod/bugs/695/
https://sourceforge.net/p/googlesyncmod/support-requests/354/, and others
- **FIX**: Invalid recurrence pattern for yearly events, see
https://sourceforge.net/p/googlesyncmod/support-requests/324/
https://sourceforge.net/p/googlesyncmod/support-requests/363/
https://sourceforge.net/p/googlesyncmod/support-requests/344/
- **IMPROVEMENT**: Swtiched to number textboxes for the months range

**Free Open Source Software, Hell Yeah!**

## Version [3.9.4] - 07.04.2015

### SVN commits

**r529 - r534**

- **FIX**: persist GoogleCalendar setting into Registry, see <https://sourceforge.net/p/googlesyncmod/bugs/685/> <https://sourceforge.net/p/googlesyncmod/bugs/684/>
- **FIX**: FIX: more spelling corrections
- **FIX**: spelling/typos corrections [bugs:#662] - UPD: nuget packages

**Free Open Source Software, Hell Yeah!**

## Version [3.9.3] - 04.04.2015

### SVN commits

**r514 - r528**

- **FIX**: fixed Google Exception when syncing appointments accepted on Google side (sent by different Organizer on Google), see <http://sourceforge.net/p/googlesyncmod/bugs/532/>
- **FIX**: not show delete conflict resoultion, if syncDelete is switched off or GoogleToOutlookOnly or OutlookToGoogleOnly
- **FIX**: fixed some issues with GoogleCalendar choice
- **FIX**: fixed some NullPointerExceptions

- **IMPROVEMENT**: Added Google Calendar Selection for appointment sync

- **IMPROVEMENT**: set culture for main-thread and SyncThread to English for english-style exception messages which are not handled by Errorhandler.cs

**Free Open Source Software, Hell Yeah!**

[3.9.3] <http://sourceforge.net/projects/googlesyncmod/files/Releases/3.9.3/SetupGCSM-3.9.3.msi/download>

## Version [3.9.2] - 27.12.2014

### SVN commits

**r511 - r513**

- **FIX**: Switched from Debugging to Release, prepared setup 3.9.2
- **FIX**: Handle AccessViolation exceptions to avoid crashes when accessing RTF Body

**Free Open Source Software, Hell Yeah!**

## Version [3.9.1] - 27.12.2014

### SVN commits

**r491 - r510**

- **FIX**: Handle Google Contact Photos wiht oAuth2 AccessToken
- **FIX**: small text changes in error dialog (added "hint message")
- **FIX**: moved client_secrets.json to Resources + added paths
- **FIX**: upgraded UnitTests and made them compilable
- **FIX**: Proxy Port was not used, because of missing exclamation mark before the null check
- **FIX**: bugfixes for Calendar sync
- **FIX**: replaced ClientLoginAuthenticator by OAuth2 Version and enabled Notes sync again
- **FIX**: removed 5 minutes minimum timespan again (doesn't make sense for 2 syncs, would make sense between changes of Outlook items, but this we cannot control
- **FIX**: Instead of deleting the registry settings, copy it from old WebGear structure ...
- **FIX**: copy error message to clipboard see [bugs:#542]

- **CHANGE**: search only .net 4.0 full profile as startup condition

- **CHANGE**: changed Auth-Class

  ```
          removed password field
          added possibility to delete user auth tokens
          changed auth folder
          changed registry settings tree name
          remove old settings-tree
  ```

- **CHANGE**: use own OAuth-Broker

  ```
          added own implementation of OAuth2-Helper class to append user (parameter: login_hint) to authorization url
          add user email to authorization uri
  ```

- **CHANGE**: removed build setting for old GoogleAPIDir

- **IMPROVEMENT**: simplified code

  ```
               rename class file - small code cleanup
  ```

- **IMPROVEMENT**: Authentication between GCSM and Google is done with OAuth2 - no password needed anymore
- **IMPROVEMENT**: changed layout and added labels for appointment fields

  ```
               set timezone before appointment sync! see [feature-requests:#112]
  ```

- **IMPROVEMENT**: setting culture for error messages to english

**Free Open Source Software, Hell Yeah!**

## Version 3.9.0

FIX: Got UnitTests running and confirmed pass results, to create setup for version 3.9.0 FIX: crash with .NET4.0 because of AccessViolationException when accessing RTFBoxy <http://sourceforge.net/p/googlesyncmod/bugs/528> FIX: Make use of Timezone settings for recurring events optional FIX: small text changes in error dialog (added "hint message") FIX: moved client_secrets.json to Resources FIX: upgraded UnitTests and made them compilable FIX: log and auth token are now written to System.Environment.SpecialFolder.ApplicationData + -NET 4.0 is now prerequisite IMPROVEMENT: added Google.Apis.Calendar.v3 and replaced v2

Contact Sync Mod, Version 3.8.6 Switched off Calandar sync, because v2 API was switched off Created last .NET 2.0 setup for version 3.8.6 (without CalendarSync

- fixed newline spelling in Error Dialog
- disable Checkbox "Run program at startup" if we can't write to hive (HKCU)
- Unload Outlook after version detection FIX: check, if Proxy settings are valid
- release outlook COM explicitly
- show Outlook Logoff in log windows
- remove old windows version detection code

Contact Sync Mod, Version 3.8.5 FIX: Handle invalid characters in syncprofile FIX: Also enable recreating appointment from Outlook to Google, if Google appointment was deleted and Outlook has multiple participants FIX: also sync 0 minutes reminder

Contact Sync Mod, Version 3.8.4 FIX: debug instead of warning, if AllDay/Start/End cannot be updated for a recurrence FIX: Don't show error messge, if appointment/contact body is the same and must not be updated

Contact Sync Mod, Version 3.8.3 Improvement: Added some info to setting errors (Google credentials and not selected folder), and added a dummy entry to the Outlook folder comboboxes to highlight, that a selection is necessary FIX: Show text, not class in Error message for recurrence FIX: Changed RTF error message to Debug FIX: Try/Catch exception when converting RTF to plain text, because some users reported memory exception since 3.8.2 release and changed error message to Debug INSTALL: added version detection for Windows 8.1 and Windows Server 2012 R2

- fixed detect of windows version
- remove "old" unmanaged calls
- use WMI to detect version

Contact Sync Mod, Version 3.8.2 IMPROVEMENT: Not overwrite RTF in Outlook contact or appointment bode FIX: recurrence exception during more than one day but not allday events are synced properly now FIX: Sensitivity can only be changed for single appointments or recurrence master

Contact Sync Mod, Version 3.8.1 FIX: sync reminder for newly created recurrence AppointmentSync IMPROVEMENT: sync private flag FIX: don't use allday property to find OriginalDate FIX: Sync deleted appointment recurrences

Contact Sync Mod, Version 3.8.0 IMPROVEMENT: Upgraded development environment from VS2010 to VS2012 and migrated setup from vdproj to wix ATTENTION: To install 3.8.0 you will have to uninstall old GCSM versions first, because the new setup (based on wix) is not compatible with the old one (based on vdproj) FIX: Save OutlookAppointment 2 times, because sometimes ComException is thrown FIX: Cleaned up some duplicate timezone entries FIX: handle Exception when permission denied for recurrences

Contact Sync Mod, Version 3.7.3 FIX: Handle error when Google contact group is not existing FIX: Handle appointments with multiple participants (ConflictResolver)

Contact Sync Mod, Version 3.7.2 FIX: don't update or delete Outlook appointments with more than 1 recipient (i.e. has been sent to participants) <https://sourceforge.net/p/googlesyncmod/support-requests/272/> FIX: Also consider changed recurrence exceptions on Google Side

Contact Sync Mod, Version 3.7.1 IMPROVEMENT: Added Timezone Combobox for Recurrent Events FIX: Fixed some pilot issues with the first appointment sync

Contact Sync Mod, Version 3.7.0 IMPROVEMENT: Added Calendar Appointments Sync

Contact Sync Mod, Version 3.6.1 FIX: Renamed automization by automation FIX: stop time, when Error is handled, to avoid multiple error message popping up

Contact Sync Mod, Version 3.6.0 IMPROVEMENT: Added icons to show syncing progress by rotating icon in notification area IMPROVEMENT: upgraded to Google Data API 2.2.0 IMPROVEMENT: linked notifyIcon.Icon to global properties' resources IMPROVEMENT: centralized all images and icon into Resources folder and replaced embedded images by link to file

Contact Sync Mod, Version 3.5.25 FIX: issue reported regarding sync folders always set to default: <https://sourceforge.net/p/googlesyncmod/bugs/436/> FIX: NullPointerException when resolving deleted GoogleNote to update again from Outlook

Contact Sync Mod, Version 3.5.24 IMPROVEMENT: Added CancelButton to cancel a running sync thread FIX: DoEvents to handle AsyncUpload of Google notes FIX: suspend timer, if user changes the time interval (to prevent running the sync instantly e.g. if removing the interval) FIX: little code cleanup FIX: add Outlook 2013 internal version number for detection FIX: removed obsolete debug-code

Contact Sync Mod, Version 3.5.23 IMPROVEMENT: Added new Icon with exclamation mark for warning/error situations FIX: show conflict in icon text and balloon, and keep conflict dialog on Top, see <http://sourceforge.net/p/googlesyncmod/support-requests/184/> FIX: Allow Outlook notes without subject (create untitled Google document) FIX: Wait up to 10 seconds until thread is alive (instead of endless loop)

Contact Sync Mod, Version 3.5.22 IMPROVEMENT: Replaced lock by Interlocked to exit sync thread if already another one is running IMPROVEMENT: fillSyncFolderItems only when needed (e.g. showing the GUI or start syncing or reset matches). IMPROVEMENT: Changed the start sync interval from 90 seconds to 5 minutes to give the PC more time to startup

Contact Sync Mod, Version 3.5.21 FIX: Fixed the issue, if Google username had an invalid character for Outlook properties <https://sourceforge.net/tracker/?func=detail&atid=1539126&aid=3598515&group_id=369321> <https://sourceforge.net/tracker/?func=detail&aid=3590035&group_id=369321&atid=1539126> FIX: Assign relationship, if no EmailDisplayName exists IMPROVEMENT: Added possibility to delete Google contact without unique property FIX: docked right splitContainer panel of ConflictResolverForm to fill full panel

Contact Sync Mod, Version 3.5.20 IMPROVEMENT: Improved INSTALL PROCESS

```
- added POSTBUILDEVENT to add version of Variable Productversion (vdproj) automatically to installer (msi) file after successful build only change the version string in the setup project and all other is done
- changed standard setup filename
```

IMPROVEMENT: added to error message to use the latest version (with url) before reporting a error to the tracker IMPROVEMENT: Added Exit-Button between hide button (Tracker ID: 3578131) FIX: Delete Google Note categories first before reassigning them (has been fixed also on Google Drive now, when updating a document, it doesn't lose the categories anymore) FIX: Updated Email Display Name

Contact Sync Mod, Version 3.5.19 IMPROVEMENT: Added Note Category sync FIX: Google Notes folder link is removed from updated note => Move note to Notes folder again after update IMPROVEMENT: added class VersionInformation (detect Outlook-Version and Operating-System-Version)

Contact Sync Mod, Version 3.5.18 FIX: added log message, if EmailDisplayName is different, because Outlook cannot set it manually FIX: switched to x86 compilation (tested with Any CPU and 64 bit, no real performance improvement), therefore x86 will be the most compatible way FIX: Preserve Email Display Name if address not changed, see also <https://sourceforge.net/tracker/index.php?func=detail&aid=3575688&group_id=369321&atid=1539129> FIX: removed Cleanup algorithm to get rid of duplicate primary phone numbers FIX: Handle unauthorized access exception when saving 'run program at startup' setting to registry, see also <https://sourceforge.net/tracker/?func=detail&aid=3560905&group_id=369321&atid=1539126> FIX: Fixed null addresses at emails

Contact Sync Mod, Version 3.5.17 FIX: applied proper tooltips to the checkboxes, see <https://sourceforge.net/tracker/?func=detail&atid=1539126&aid=3559759&group_id=369321> FIX: UI Spelling and Grammar Corrections - ID: 3559753 FIX: fixed problem when saving Google Photo, see <https://sourceforge.net/tracker/?func=detail&aid=3555588&group_id=369321&atid=1539126>

Contact Sync Mod, Version 3.5.16 FIX: fixed bug when deleting a contact on GoogleSide (Precondition failed error) FIX: fixed some typos and label sizes in ConflictResolverForm FIX: Also handle InvalidCastException when loggin into Outlook IMPROVEMENT: changed some variable declarations to var FIX: Skip empty OutlookNote to avoid Nullpointer Reference Exception FIX: fixed IM sync, not to add the address again and again, until the storage of this field exceeds on Google side FIX: fixed saving contacts and notes folder to registry, if empty before

Contact Sync Mod, Version 3.5.15 FIX: increased TimeTolerance to 120 seconds to avoid resync after ResetMatches FIX: added UseFileAs feature also for updating existing contacts IMPROVEMENT: applied "UseFileAs" setting also for syncing from Google to Outlook (to allow Outlook to choose FileAs as configured) IMPROVEMENT: replaced radiobuttons rbUseFileAs and rbUseFullName by Checkbox chkUseFileAs and moved it from bottom to the settings groupBox

Contact Sync Mod, Version 3.5.14 FIX: NullPointerException when syncing notes, see <https://sourceforge.net/tracker/index.php?func=detail&aid=3522539&group_id=369321&atid=1539126> IMPROVEMENT: Added setting to choose between Outlook FileAs and FullName

Contact Sync Mod, Version 3.5.13 IMPROVEMENT: added tooltips to Username and Password if input is wrong IMPROVEMENT: put contacts and notes folder combobox in different lines to enable resizing them Improvement: Migrated to Google Data API 2.0 Imporvement: switched to ResumableUploader for GoogleNotes FIX: Changed layer order of checkboxes to overcome hiding them, if Windows is showing a bigger font

Contact Sync Mod, Version 3.5.12 IMPROVEMENT: Implemented GUI to match Duplicates and added feature to keep both (Google and Outlook entry) FIX: Only show warning, if an OutlookFolder couldn't be opened and try to open next one

Contact Sync Mod, Version 3.5.11 FIX: Also create Outlook Contact and Note items in the selected folder (not default folder)

Contact Sync Mod, Version 3.5.10 FIX: Only check log file size, if log file size already exists

Contact Sync Mod, Version 3.5.9 IMPROVEMENT: create new logfile, once 1MB has been exceeded (move to backup before) Improvement: Added ConflictResolutions to perform selected actions for all following itmes IMPROVEMENT: Enable the user to configure multipole sync profiles, e.g. to sync with multiple gmail accounts IMPROVEMENT: Enable the user to choose Outlook folder IMPROVEMENT: Added language sync FIX: Remove Google Note directly from root folder IMPROVEMENT: No ErrorHandle when neither notes nor contacts are selected ==> Show BalloonTooltip and Form instead Improvement: Added ComException special handling for not reachable RPC, e.g. if Outlook was closed during sync Improvement: Added SwitchTimer to Unlock PC message FIX: Improved error handling, especially when invalid credentials=> show the settings form Improvement: handle Standby/Hibernate and Resume windows messages to suspend timer for 90 seconds after resume

Contact Sync Mod, Version 3.5.8 FIX: validation mask of proxy user name (by #3469442) FIX: handled OleAut Date exceptions when updating birthday IMPROVEMENT: open Settings GUI of running GCSM process when starting new instance (instead of error message, that a process is already running) FIX: validation mask of proxy uri (by #3458192) IMPROVEMENT: ResetMatch when deleting an entry (to avoid deleting it again, if restored from Outlook recycle bin)

Contact Sync Mod, Version 3.5.7 IMPROVEMENT: made OutlookApplication and Namespace static IMPROVEMENT: added balloon after first run, see <https://sourceforge.net/tracker/?func=detail&aid=3429308&group_id=369321&atid=1539126> FIX: Delete temporary note file before creating a new one FIX: Reset OutlookGoogleNoteId after note has been deleted on Google side before recreated by Upload (new GoogleNoteId) FIX: Set bypass proxy local resource in new proxy mask FIX: set for use default credentials for auth. in new proxy mask

Contact Sync Mod, Version 3.5.6 IMPROVEMENT: added proxy config mask and proxy authentication (in addition to use App.config) IMPROVEMENT: finished Notes sync feature IMPROVEMENT: Switched to new Google API 1.9 (Previous: 1.8) FIX: Added CreateOutlookInstance to OutlookNameSpace property, to avoid NullReferenceExceptions FIX: Removed characters not allowed for Outlook user property names: []_# FIX: handled exception when updating Birthday and anniversary with invalid date, see <https://sourceforge.net/tracker/?func=detail&aid=3397921&group_id=369321&atid=1>

Contact Sync Mod, Version 3.5.5 FIX: set _sync.SyncContacts properly when resetting matches (fixes <https://sourceforge.net/tracker/index.php?func=detail&aid=3403819&group_id=369321&atid=1539126>)

Contact Sync Mod, Version 3.5.4 IMPROVEMENT: added pdb file to installation to get some more information, when users report bugs IMPROVEMENT: Added also email to not require FullName IMPROVEMENT: Added company as unique property, if FullName is emptyFullName

See also Feature Request <https://sourceforge.net/tracker/index.php?func=detail&aid=3297935&group_id=369321&atid=1539126> FIX: handled exception when updating Birthday and anniversary with invalid date, see <https://sourceforge.net/tracker/?func=detail&aid=3397921&group_id=369321&atid=1539126> FIX: Handle Nullpointerexception when Release Marshall Objects at GetOutlookItems, maybe this helps to fix the Nullpointer Exceptions in LoadOutlookContacts

Contact Sync Mod, Version 3.5.3 Improvement: Upgraded to Google Data API 1.8 FIX: Handle Nullpointerexception when Release Marshall Objects

Contact Sync Mod, Version 3.5.1

FIX: Handle AccessViolation Exception when trying to get Email address from Exchange Email

Contact Sync Mod, Version 3.5

FIX: Moved NotificationReceived to constructor to not handle this event redundantly FIX: moved assert of TestSyncPhoto above the UpdateContact line FIX: Added log message when skipping a faulty Outlook Contact FIX: fixed number of current match (i not i+1) because of 1 based array Fix: set SyncDelete at every SyncStart to avoid "Skipped deletion" warnings, though Sync Deletion checkbox was checked Improvement: Support large Exchange contact lists, get SMTP email when Exchange returns X500 addresses, use running Outlook instance if present.

CHANGE 1: Support a large number of contacts on Exchange server without hitting the policy limitation of max number of contacts that can be processed simultaneously.

CHANGE 2: Enhancement request 3156687: Properly get the SMTP email address of Exchange contacts when Exchange returns X500 addresses.

CHANGE 3: Try to contact a running Outlook application before trying to launch a new one. Should make the program work in any situation, whether Outlook is running or not.

OTHER SMALL FIXES:

- Never re-throw an exception using "throw ex". Just use "throw". (preserves stack trace)
- Handle an invalid photo on a Google contact (skip the photo).

IMPROVEMENT: added EnableLaunchApplication to start GOContactSyncMod after installation as PostBuildEvent Improvement: added progress notifications (which contact is currently syncing or matching) Improvement: Sync also State and PostOfficeBox, see Tracker item <https://sourceforge.net/tracker/?func=detail&aid=3276467&group_id=369321&atid=1539126> Improvement: Avoid MatchContacts when just resetting matches (Performance improvement)

[3.9.1]: http://sourceforge.net/projects/googlesyncmod/files/Releases/3.9.1/SetupGCSM-3.9.1.msi/download
[3.9.2]: http://sourceforge.net/projects/googlesyncmod/files/Releases/3.9.2/SetupGCSM-3.9.2.msi/download