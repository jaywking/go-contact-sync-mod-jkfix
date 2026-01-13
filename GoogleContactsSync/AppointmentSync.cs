using Google.Apis.Calendar.v3.Data;
using NodaTime;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal static class AppointmentSync
    {
        private const string dateFormat = "yyyyMMdd";
        private const string timeFormat = "HHmmss";
        internal static readonly DateTime outlookDateInvalid = new DateTime(4501, 1, 1);
        internal static readonly DateTime outlookDateMax = new DateTime(4500, 12, 31);
        private const string RRULE = "RRULE";
        private const string FREQ = "FREQ";
        private const string DAILY = "DAILY";
        private const string WEEKLY = "WEEKLY";
        private const string MONTHLY = "MONTHLY";
        private const string YEARLY = "YEARLY";
        private const string BYMONTH = "BYMONTH";
        private const string BYMONTHDAY = "BYMONTHDAY";
        private const string BYDAY = "BYDAY";
        private const string BYSETPOS = "BYSETPOS";
        private const string INTERVAL = "INTERVAL";
        private const string COUNT = "COUNT";
        private const string UNTIL = "UNTIL";
        private const string MO = "MO";
        private const string TU = "TU";
        private const string WE = "WE";
        private const string TH = "TH";
        private const string FR = "FR";
        private const string SA = "SA";
        private const string SU = "SU";
        internal const string PRIVATE = "private";
        private const string CONFIDENTIAL = "confidential";
        private const string OPAQUE = "opaque";
        private const string CONFIRMED = "confirmed";
        private const string TRANSPARENT = "transparent";
        private const string TENTATIVE = "tentative";

        // This will return the Windows zone that matches the IANA zone, if one exists.
        internal static string IanaToWindows(string ianaZoneId)
        {
            var utcZones = new[] { "Etc/UTC", "Etc/UCT", "Etc/GMT" };
            if (utcZones.Contains(ianaZoneId, StringComparer.Ordinal))
            {
                return "UTC";
            }

            var tzdbSource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;

            // resolve any link, since the CLDR doesn't necessarily use canonical IDs
            var links = tzdbSource.CanonicalIdMap
                .Where(x => x.Value.Equals(ianaZoneId, StringComparison.Ordinal))
                .Select(x => x.Key);

            // resolve canonical zones, and include original zone as well
            var possibleZones = tzdbSource.CanonicalIdMap.ContainsKey(ianaZoneId)
                ? links.Concat(new[] { tzdbSource.CanonicalIdMap[ianaZoneId], ianaZoneId })
                : links;

            // map the windows zone
            var mappings = tzdbSource.WindowsMapping.MapZones;
            var item = mappings.FirstOrDefault(x => x.TzdbIds.Any(possibleZones.Contains));
            return item?.WindowsId;
        }

        // This will return the "primary" IANA zone that matches the given windows zone.
        // If the primary zone is a link, it then resolves it to the canonical ID.
        private static string WindowsToIana(string windowsZoneId)
        {
            // Avoid UTC being mapped to Etc/GMT, which is the mapping in CLDR
            if (windowsZoneId == "UTC")
            {
                return "Etc/UTC";
            }
            var source = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;
            // If there's no such mapping, result will be null.
            source.WindowsMapping.PrimaryMapping.TryGetValue(windowsZoneId, out var result);
            // Canonicalize
            if (result != null)
            {
                result = source.CanonicalIdMap[result];
            }
            return result;
        }

        public static DateTime LocaltoUTC(DateTime dateTime, string IanaZone)
        {
            var localDateTime = LocalDateTime.FromDateTime(dateTime);
            var usersTimezone = DateTimeZoneProviders.Tzdb[IanaZone];
            var zonedDbDateTime = usersTimezone.AtLeniently(localDateTime);
            return zonedDbDateTime.ToDateTimeUtc();
        }

        private static void UpdateStartDate(Outlook.AppointmentItem master, Event slave)
        {
            if (slave.Start == null)
                slave.Start = new EventDateTime();

            if (master.AllDayEvent == true)
            {
                slave.Start.Date = master.Start.ToString("yyyy-MM-dd");
                slave.Start.DateTimeDateTimeOffset = null;
            }
            else
            {
                //Outlook always has TZ set, even if TZ is the same as default one
                //Google could have TZ empty, if it is equal to default one
                var google_start_tz = string.Empty;

                // StartTimeZone was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                Outlook.TimeZone outlook_start_tz = null;
                try
                {
                    outlook_start_tz = master.StartTimeZone;
                    if (outlook_start_tz != null)
                    {
                        if (!string.IsNullOrEmpty(outlook_start_tz.ID))
                        {
                            google_start_tz = WindowsToIana(outlook_start_tz.ID);
                        }
                    }
                }
                catch (AccessViolationException)
                {
                }
                finally
                {
                    if (outlook_start_tz != null)
                    {
                        Marshal.ReleaseComObject(outlook_start_tz);
                    }
                }

                slave.Start.Date = null;

                if (string.IsNullOrEmpty(google_start_tz))
                {
                    slave.Start.DateTimeDateTimeOffset = new DateTimeOffset(master.Start);
                }
                else
                {
                    //todo (obelix30), workaround for https://github.com/google/google-api-dotnet-client/issues/853
                    var zone = DateTimeZoneProviders.Tzdb[google_start_tz];
                    var start_local = LocalDateTime.FromDateTime(master.StartInStartTimeZone);
                    var start_zoned = start_local.InZoneLeniently(zone);
                    var start_utc = start_zoned.ToDateTimeUtc();
                    slave.Start.DateTimeDateTimeOffset = new DateTimeOffset(start_utc);
                    if (google_start_tz != Synchronizer.SyncAppointmentsGoogleTimeZone)
                    {
                        slave.Start.TimeZone = google_start_tz;
                    }
                }
            }
        }

        private static void UpdateEndDate(Outlook.AppointmentItem master, Event slave)
        {
            if (slave.End == null)
                slave.End = new EventDateTime();

            if (master.AllDayEvent == true)
            {
                slave.End.Date = master.End.ToString("yyyy-MM-dd");
                slave.End.DateTimeDateTimeOffset = null;
            }
            else
            {
                //Outlook always has TZ set, even if TZ is the same as default one
                //Google could have TZ empty, if it is equal to default one
                var google_end_tz = string.Empty;

                // EndTimeZone was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                Outlook.TimeZone outlook_end_tz = null;
                try
                {
                    outlook_end_tz = master.EndTimeZone;
                    if (outlook_end_tz != null)
                    {
                        if (!string.IsNullOrEmpty(outlook_end_tz.ID))
                        {
                            google_end_tz = WindowsToIana(outlook_end_tz.ID);
                        }
                    }
                }
                catch (AccessViolationException)
                {
                }
                finally
                {
                    if (outlook_end_tz != null)
                    {
                        Marshal.ReleaseComObject(outlook_end_tz);
                    }
                }

                slave.End.Date = null;

                if (string.IsNullOrEmpty(google_end_tz))
                {
                    slave.End.DateTimeDateTimeOffset = new DateTimeOffset(master.End);
                }
                else
                {
                    //todo (obelix30), workaround for https://github.com/google/google-api-dotnet-client/issues/853
                    var zone = DateTimeZoneProviders.Tzdb[google_end_tz];
                    var end_local = LocalDateTime.FromDateTime(master.EndInEndTimeZone);
                    var end_zoned = end_local.InZoneLeniently(zone);
                    var end_utc = end_zoned.ToDateTimeUtc();
                    slave.End.DateTimeDateTimeOffset = new DateTimeOffset(end_utc);
                    if (google_end_tz != Synchronizer.SyncAppointmentsGoogleTimeZone)
                    {
                        slave.End.TimeZone = google_end_tz;
                    }
                }
            }
        }

        /// <summary>
        /// Updates Outlook appointments (calendar) to Google Calendar
        /// </summary>
        public static void UpdateAppointment(Outlook.AppointmentItem master, Event slave, bool skip_recurrence_sync = false)
        {
            slave.Summary = master.Subject;

            if (Synchronizer.SyncAppointmentsPrivate && AppointmentsMatcher.RecipientsCount(master) > 1) //If Outlook Appointments shall be private and multiple outlook recipients found, don't copy the appointment content/body (e.g. to hide business appointments to Google or private Google appointments to business)
                slave.Description = string.IsNullOrEmpty(master.Body) ? master.Body : "!!!GCSM: Content not copied over from Outlook, because Appointments private on other side setting and multiple outlook recipients, please check your Outlook appointment to view the content of this appointment!!!";            
            else
                slave.Description = master.Body;

            switch (master.BusyStatus)
            {
                case Outlook.OlBusyStatus.olBusy:
                    slave.Transparency = OPAQUE;
                    slave.Status = CONFIRMED;
                    break;
                case Outlook.OlBusyStatus.olTentative:
                    slave.Transparency = TRANSPARENT;
                    slave.Status = TENTATIVE;
                    break;
                case Outlook.OlBusyStatus.olOutOfOffice:
                    slave.Transparency = OPAQUE;
                    slave.Status = TENTATIVE;
                    break;
                //ToDo: case Outlook.OlBusyStatus.olWorkingElsewhere:
                case Outlook.OlBusyStatus.olFree:
                default:
                    slave.Status = CONFIRMED;
                    slave.Transparency = TRANSPARENT;
                    break;
            }

            slave.Location = master.Location;


            UpdateStartDate(master, slave);
            UpdateEndDate(master, slave);

            if (Synchronizer.SyncReminders)
                UpdateReminders(master, slave);

            if (!skip_recurrence_sync)
            {
                UpdateRecurrence(master, slave);
            }

            //if (slave.Recurrence == null || slave.Recurrence.Count == 0)
            //{
                var oldVisibility = slave.Visibility;

                switch (master.Sensitivity)
                {
                    case Outlook.OlSensitivity.olConfidential: //ToDo, currently not supported by Google Web App GUI and Outlook 2010: slave.EventVisibility = Google.GData.Calendar.Event.Visibility.CONFIDENTIAL; break;#
                    case Outlook.OlSensitivity.olPersonal: //ToDo, currently not supported by Google Web App GUI and Outlook 2010: slave.EventVisibility = Google.GData.Calendar.Event.Visibility.CONFIDENTIAL; break;
                    case Outlook.OlSensitivity.olPrivate: slave.Visibility = PRIVATE; break;
                    default: slave.Visibility = "default"; break;
                }

                if (Synchronizer.SyncAppointmentsPrivate) //If Appointments shall be private on the other side, set it (e.g. to hide business appointments to Google or private Google appointments to business)
                {
                    if (oldVisibility == PRIVATE || //Don'T overwrite privacy setting, if it was set already to private (e.g. to avoid resyncing here the wrong privacy)
                        string.IsNullOrEmpty(AppointmentPropertiesUtils.GetOutlookGoogleId(master))) //or if newly created on Google side
                        slave.Visibility = PRIVATE;
                    else
                        slave.Visibility = oldVisibility;
                }
            //}
        }

        private static void UpdateReminders(Outlook.AppointmentItem master, Event slave)
        {
            if (slave.Reminders == null)
            {
                slave.Reminders = new Event.RemindersData
                {
                    Overrides = new List<EventReminder>()
                };
            }
            else
            {
                if (slave.Reminders.Overrides != null)
                {
                    slave.Reminders.Overrides.Clear();
                }
                slave.Reminders.UseDefault = false;
            }

            slave.Reminders.UseDefault = false;
            if (slave.Reminders.Overrides != null)
            {
                slave.Reminders.Overrides.Clear();
            }
            else
            {
                slave.Reminders.Overrides = new List<EventReminder>();
            }

            if (master.ReminderSet && (master.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster || master.Start > DateTime.Now || Synchronizer.IncludePastReminders))
            {
                var reminder = new EventReminder
                {
                    Minutes = master.ReminderMinutesBeforeStart
                };
                if (reminder.Minutes > 40300)
                {
                    //ToDo: Check real limit, currently 40300
                    Log.Warning("Reminder Minutes to big (" + reminder.Minutes + "), set to maximum of 40300 minutes for appointment: " + master.ToLogString());
                    reminder.Minutes = 40300;
                }
                reminder.Method = "popup";
                slave.Reminders.Overrides.Add(reminder);
            }
            
        }

        private static void UpdateBody(Event master, Outlook.AppointmentItem slave)
        {
            try
            {
                string nonRTF;

                // RTFBody was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                try
                {
                    nonRTF = slave.Body == null ? string.Empty : slave.RTFBody != null ? Utilities.ConvertToText(slave.RTFBody as byte[]) : string.Empty;
                }
                catch (AccessViolationException)
                {
                    nonRTF = slave.Body ?? string.Empty;
                }

                if (!nonRTF.Equals(master.Description))
                {
                    if (string.IsNullOrEmpty(nonRTF) || nonRTF.Equals(slave.Body))
                    {  //only update, if RTF text is same as plain text and is different between master and slave
                        slave.Body = master.Description;
                    }
                    else
                    {
                        if (Synchronizer.SyncAppointmentsForceRTF)
                        {
                            slave.Body = master.Description;
                        }
                        else
                        {
                            Log.Warning("Outlook appointment notes body not updated, because it is RTF, otherwise it will overwrite it by plain text: " + slave.ToLogString());
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Log.Debug(e, "Error when converting RTF to plain text, updating Google Appointment '" + slave.ToLogString() + "' notes to Outlook without RTF check: " + e.Message);
                try
                {
                    if (slave.Body != master.Description)
                        slave.Body = master.Description;
                }
                catch (Exception e2)
                {
                    Log.Error(e2, "Error when updating Google Appointment '" + slave.ToLogString() + "' notes to Outlook without RTF check: " + e.Message);
                }
            }
        }

        private static void UpdateStartTime(Event master, Outlook.AppointmentItem slave)
        {
            // before setting times in Outlook, set correct time zone
            // StartTimeZone was introduced in later version of Outlook
            // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
            Outlook.Application app;
            Outlook.TimeZones tz = null;
            try
            {
                if (master.Start.TimeZone == null)
                {
                    if (Synchronizer.MappingBetweenTimeZonesRequired)
                    {
                        var outlook_tz = IanaToWindows(Synchronizer.SyncAppointmentsGoogleTimeZone);
                        app = Synchronizer.OutlookApplication;
                        tz = app.TimeZones;
                        slave.StartTimeZone = tz[outlook_tz];
                    }
                }
                else
                {
                    var outlook_tz = IanaToWindows(master.Start.TimeZone);
                    app = Synchronizer.OutlookApplication;
                    tz = app.TimeZones;
                    slave.StartTimeZone = tz[outlook_tz];
                }
            }
            catch (AccessViolationException)
            {
            }
            finally
            {
                if (tz != null)
                {
                    Marshal.ReleaseComObject(tz);
                }
            }

            //master.Start.DateTimeDateTimeOffset is specified in Google calendar default time zone
            var startUTC = LocaltoUTC(master.Start.DateTimeDateTimeOffset.Value.DateTime, Synchronizer.SyncAppointmentsGoogleTimeZone);

            // StartUTC was introduced in later version of Outlook
            // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
            try
            {
                if (slave.StartUTC != startUTC)
                {
                    slave.StartUTC = startUTC;
                }
            }
            catch (AccessViolationException)
            {
                if (slave.Start != master.Start.DateTimeDateTimeOffset.Value.DateTime)
                {
                    slave.Start = master.Start.DateTimeDateTimeOffset.Value.DateTime;
                }
            }
            catch (COMException)
            {
                if (slave.Start != master.Start.DateTimeDateTimeOffset.Value.DateTime)
                {
                    slave.Start = master.Start.DateTimeDateTimeOffset.Value.DateTime;
                }
            }
        }

        private static void UpdateEndTime(Event master, Outlook.AppointmentItem slave)
        {
            // before setting times in Outlook, set correct time zone
            // EndTimeZone was introduced in later version of Outlook
            // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
            Outlook.Application app;
            Outlook.TimeZones tz = null;
            try
            {
                if (master.End.TimeZone == null)
                {
                    if (Synchronizer.MappingBetweenTimeZonesRequired)
                    {
                        var outlook_tz = IanaToWindows(Synchronizer.SyncAppointmentsGoogleTimeZone);
                        app = Synchronizer.OutlookApplication;
                        tz = app.TimeZones;
                        slave.EndTimeZone = tz[outlook_tz];
                    }
                }
                else
                {
                    var outlook_tz = IanaToWindows(master.End.TimeZone);
                    app = Synchronizer.OutlookApplication;
                    tz = app.TimeZones;
                    slave.EndTimeZone = tz[outlook_tz];
                }
            }
            catch (AccessViolationException)
            {
            }
            finally
            {
                if (tz != null)
                {
                    Marshal.ReleaseComObject(tz);
                }
            }

            //master.End.DateTimeDateTimeOffset is specified in Google calendar default time zone
            var endUTC = LocaltoUTC(master.End.DateTimeDateTimeOffset.Value.DateTime, Synchronizer.SyncAppointmentsGoogleTimeZone);

            // EndUTC was introduced in later version of Outlook
            // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
            try
            {
                if (slave.EndUTC != endUTC)
                {
                    slave.EndUTC = endUTC;
                }
            }
            catch (AccessViolationException)
            {
                if (slave.End != master.End.DateTimeDateTimeOffset.Value.DateTime)
                {
                    slave.End = master.End.DateTimeDateTimeOffset.Value.DateTime;
                }
            }
            catch (COMException)
            {
                if (slave.End != master.End.DateTimeDateTimeOffset.Value.DateTime)
                {
                    slave.End = master.End.DateTimeDateTimeOffset.Value.DateTime;
                }
            }
        }

        /// <summary>
        /// Updates Outlook appointments (calendar) to Google Calendar
        /// </summary>
        public static bool UpdateAppointment(Event master, Outlook.AppointmentItem slave)
        {
            slave.Subject = master.Summary;

            UpdateBody(master, slave);

            slave.Location = master.Location;

            try
            {
                if (master.Start != null && slave.AllDayEvent == string.IsNullOrEmpty(master.Start.Date))
                {
                    slave.AllDayEvent = !string.IsNullOrEmpty(master.Start.Date);
                }

                if (master.Start != null && !string.IsNullOrEmpty(master.Start.Date))
                {
                    var d = DateTime.Parse(master.Start.Date);

                    if (d != slave.Start)
                    {
                        slave.Start = d;
                    }
                }
                else if (master.Start != null && master.Start.DateTimeDateTimeOffset != null)
                {
                    UpdateStartTime(master, slave);
                }
                if (master.End != null && !string.IsNullOrEmpty(master.End.Date))
                {
                    var d = DateTime.Parse(master.End.Date);

                    if (d != slave.End)
                    {
                        slave.End = d;
                    }
                }
                else if (master.End != null && master.End.DateTimeDateTimeOffset != null)
                {
                    UpdateEndTime(master, slave);
                }
            }
            catch (Exception ex)
            {
                //if (slave.IsRecurring)
                //{
                    Log.Debug("Error updating event's AllDay/Start/End: " + master.ToLogString() + ": " + ex.Message);
                //}
                //else
                //{
                //    Log.Warning("Error updating event's AllDay/Start/End: " + master.ToLogString() + ": " + ex.Message);
                //}

                try
                {
                    //If updating End fails for recurring event, update at least duration
                    var start = master.Start.DateTimeDateTimeOffset.HasValue ? master.Start.DateTimeDateTimeOffset.Value : DateTime.Parse(master.Start.Date);
                    var end = master.End.DateTimeDateTimeOffset.HasValue ? master.End.DateTimeDateTimeOffset.Value : DateTime.Parse(master.End.Date);
                    TimeSpan span = end - start;
                    if (slave.Duration != span.Minutes)
                        slave.Duration = span.Minutes;
                }
                catch (Exception ex2)
                {
                    Log.Warning("Error updating event's Duration: " + master.ToLogString() + ": " + ex.Message + " / " + ex2.Message);
                    Log.Debug(ex, "Exception");
                    Log.Debug(ex2, "Exception");
                    master.ToDebugLog();
                    slave.ToDebugLog();
                    //return false;

                }
            }


            try
            {
                if (master.Status.Equals(CONFIRMED) && (master.Transparency == null || master.Transparency.Equals(OPAQUE)))
                {
                    slave.BusyStatus = Outlook.OlBusyStatus.olBusy;
                }
                else if ((master.Status.Equals(CONFIRMED) && master.Transparency.Equals(TRANSPARENT)) || master.Status.Equals("cancelled"))
                {
                    slave.BusyStatus = Outlook.OlBusyStatus.olFree;
                }
                else
                {
                    slave.BusyStatus = master.Status.Equals(TENTATIVE) && (master.Transparency == null || master.Transparency.Equals(OPAQUE))
                        ? Outlook.OlBusyStatus.olOutOfOffice
                        : master.Status.Equals(TENTATIVE) ? Outlook.OlBusyStatus.olTentative : Outlook.OlBusyStatus.olWorkingElsewhere;
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                master.ToDebugLog();
                slave.ToDebugLog();
                return false;
            }

            if (Synchronizer.SyncReminders)
                slave.ReminderSet = false;
            if (Synchronizer.SyncReminders && (slave.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster || slave.Start > DateTime.Now || Synchronizer.IncludePastReminders))
            {
                if (master.Reminders != null)
                {
                    if (master.Reminders.UseDefault != null)
                    {
                        slave.ReminderSet = master.Reminders.UseDefault.Value;
                    }

                    if (master.Reminders.Overrides != null)
                    {
                        foreach (var reminder in master.Reminders.Overrides)
                        {
                            if (reminder.Method == "popup" && reminder.Minutes != null)
                            {
                                slave.ReminderSet = true;
                                slave.ReminderMinutesBeforeStart = reminder.Minutes.Value;
                            }
                        }
                    }
                }
            }

            if (!UpdateRecurrence(master, slave))
            {
                return false;
            }

            try
            {
                //Sensivity update is only allowed for single appointments or the master
                if (!slave.IsRecurring || slave.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster)
                {
                    var oldSensitivity = slave.Sensitivity;
                    switch (master.Visibility)
                    {
                        case CONFIDENTIAL:
                        case PRIVATE: slave.Sensitivity = Outlook.OlSensitivity.olPrivate; break;
                        default: slave.Sensitivity = Outlook.OlSensitivity.olNormal; break;
                    }

                    if (Synchronizer.SyncAppointmentsPrivate) //If Appointments shall be private on the other side, set it when newly created (e.g. to hide business appointments to Google or private Google appointments to business)
                    {
                        if (oldSensitivity == Outlook.OlSensitivity.olPrivate || //Don'T overwrite privacy setting, if it was set already to private (e.g. to avoid resyncing here the wrong privacy)
                         string.IsNullOrEmpty(AppointmentPropertiesUtils.GetGoogleOutlookId(master))) //or if newly created on Outlook side
                            slave.Sensitivity = Outlook.OlSensitivity.olPrivate;
                        else if (slave.Sensitivity != oldSensitivity)
                            slave.Sensitivity = oldSensitivity;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                slave.ToDebugLog();
                return false;
            }

            return true;
        }

        private static void UpdateRecurrence(Outlook.AppointmentItem master, Event slave)
        {
            if (IsSameRecurrence(slave, master))
            {
                return;
            }

            Outlook.RecurrencePattern rp = null;

            try
            {
                if (!master.IsRecurring)
                {
                    if (slave.Recurrence != null)
                    {
                        slave.Recurrence = null;
                    }

                    return;
                }

                if (master.RecurrenceState != Outlook.OlRecurrenceState.olApptMaster)
                {
                    return;
                }

                rp = master.GetRecurrence();

                var slaveRecurrence = string.Empty;

                // StartTimeZone was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                Outlook.TimeZone stz = null;
                try
                {
                    stz = master.StartTimeZone;
                    if (stz != null && !string.IsNullOrEmpty(stz.ID))
                    {
                        var google_tz = WindowsToIana(stz.ID);
                        slave.Start.TimeZone = google_tz;
                    }
                }
                catch (AccessViolationException)
                {
                }
                finally
                {
                    if (stz != null)
                    {
                        Marshal.ReleaseComObject(stz);
                    }
                }

                // EndTimeZone was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                Outlook.TimeZone etz = null;
                try
                {
                    etz = master.EndTimeZone;
                    if (etz != null && !string.IsNullOrEmpty(etz.ID))
                    {
                        var google_tz = WindowsToIana(etz.ID);
                        slave.End.TimeZone = google_tz;
                    }
                }
                catch (AccessViolationException)
                {
                }
                finally
                {
                    if (etz != null)
                    {
                        Marshal.ReleaseComObject(etz);
                    }
                }

                if (slave.Recurrence == null)
                {
                    slave.Recurrence = new List<string>();
                }
                else
                {
                    slave.Recurrence.Clear();
                }

                slaveRecurrence = RRULE + ":" + FREQ + "=";
                switch (rp.RecurrenceType)
                {
                    case Outlook.OlRecurrenceType.olRecursDaily: slaveRecurrence += DAILY; break;
                    case Outlook.OlRecurrenceType.olRecursWeekly: slaveRecurrence += WEEKLY; break;
                    case Outlook.OlRecurrenceType.olRecursMonthly:
                    case Outlook.OlRecurrenceType.olRecursMonthNth: slaveRecurrence += MONTHLY; break;
                    case Outlook.OlRecurrenceType.olRecursYearly:
                    case Outlook.OlRecurrenceType.olRecursYearNth: slaveRecurrence += YEARLY; break;
                    default: throw new NotSupportedException("RecurrenceType not supported by Google: " + rp.RecurrenceType);
                }

                var byDay = string.Empty;
                if ((rp.DayOfWeekMask & Outlook.OlDaysOfWeek.olMonday) == Outlook.OlDaysOfWeek.olMonday)
                {
                    byDay = MO;
                }

                if ((rp.DayOfWeekMask & Outlook.OlDaysOfWeek.olTuesday) == Outlook.OlDaysOfWeek.olTuesday)
                {
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + TU;
                }

                if ((rp.DayOfWeekMask & Outlook.OlDaysOfWeek.olWednesday) == Outlook.OlDaysOfWeek.olWednesday)
                {
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + WE;
                }

                if ((rp.DayOfWeekMask & Outlook.OlDaysOfWeek.olThursday) == Outlook.OlDaysOfWeek.olThursday)
                {
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + TH;
                }

                if ((rp.DayOfWeekMask & Outlook.OlDaysOfWeek.olFriday) == Outlook.OlDaysOfWeek.olFriday)
                {
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + FR;
                }

                if ((rp.DayOfWeekMask & Outlook.OlDaysOfWeek.olSaturday) == Outlook.OlDaysOfWeek.olSaturday)
                {
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + SA;
                }

                if ((rp.DayOfWeekMask & Outlook.OlDaysOfWeek.olSunday) == Outlook.OlDaysOfWeek.olSunday)
                {
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + SU;
                }

                if (!string.IsNullOrEmpty(byDay))
                {
                    if (rp.Instance != 0)
                    {
                        if (rp.Instance >= 1 && rp.Instance <= 4)
                        {
                            byDay = rp.Instance + byDay;
                        }
                        else if (rp.Instance == 5)
                        {
                            slaveRecurrence += ";" + BYSETPOS + "=-1";
                        }
                        else
                        {
                            throw new NotSupportedException("Outlook Appointment Instances 1-4 and 5 (last) are allowed but was: " + rp.Instance);
                        }
                    }
                    slaveRecurrence += ";" + BYDAY + "=" + byDay;
                }

                if (rp.DayOfMonth != 0)
                {
                    slaveRecurrence += ";" + BYMONTHDAY + "=" + rp.DayOfMonth;
                }

                if (rp.MonthOfYear != 0)
                {
                    slaveRecurrence += ";" + BYMONTH + "=" + rp.MonthOfYear;
                }

                if ((rp.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearly &&
                    rp.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearNth &&
                    rp.Interval > 1) ||
                    rp.Interval > 12)
                {
                    if (rp.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearly &&
                        rp.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearNth)
                    {
                        slaveRecurrence += ";" + INTERVAL + "=" + rp.Interval;
                    }
                    else
                    {
                        slaveRecurrence += ";" + INTERVAL + "=" + (rp.Interval / 12);
                    }
                }

                if (rp.PatternEndDate.Date != outlookDateInvalid &&
                    rp.PatternEndDate.Date != outlookDateMax)
                {
                    slaveRecurrence += ";" + UNTIL + "=" + rp.PatternEndDate.Date.AddDays(master.AllDayEvent ? 0 : 1).ToString(dateFormat);
                }

                slave.Recurrence.Add(slaveRecurrence);
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                Log.Error($"Error updating Google event: {slave.ToLogString()}: {ex.Message}");
            }
            finally
            {
                if (rp != null)
                {
                    Marshal.ReleaseComObject(rp);
                }
            }
        }

        public static bool IsSameRecurrence(Event ga, Outlook.AppointmentItem oa)
        {
            var gr = ga.Recurrence;

            if (!oa.IsRecurring)
            {
                if ((gr == null) || (gr.Count < 1))
                {
                    return true;
                }
                return false;
            }

            if ((gr == null) || (gr.Count < 1))
            {
                return false;
            }

            Outlook.RecurrencePattern or = null;
            try
            {
                or = oa.GetRecurrence();

                if (ga.Start != null)
                {
                    if (!ga.Start.DateTimeDateTimeOffset.HasValue)
                    {
                        if (!string.IsNullOrEmpty(ga.Start.Date))
                        {
                            if (or.StartTime.Hour != 0 || or.StartTime.Minute != 0 || or.StartTime.Second != 0)
                            {
                                return false;
                            }
                            if (!oa.AllDayEvent)
                            {
                                return false;
                            }
                        }
                    }
                    else
                    {
                        if (or.StartTime.TimeOfDay != ga.Start.DateTimeDateTimeOffset.Value.TimeOfDay)
                        {
                            return false;
                        }
                        if (or.PatternStartDate.Date != ga.Start.DateTimeDateTimeOffset.Value.Date)
                        {
                            return false;
                        }
                    }
                }

                if (ga.End != null)
                {
                    if (!ga.End.DateTimeDateTimeOffset.HasValue)
                    {
                        if (!string.IsNullOrEmpty(ga.End.Date))
                        {
                            if (or.EndTime.Hour != 0 || or.EndTime.Minute != 0 || or.EndTime.Second != 0)
                            {
                                return false;
                            }
                            if (!oa.AllDayEvent)
                            {
                                return false;
                            }
                        }
                    }
                    else
                    {
                        if (or.EndTime.TimeOfDay != ga.End.DateTimeDateTimeOffset.Value.TimeOfDay)
                        {
                            return false;
                        }
                    }
                }

                foreach (var pattern in gr)
                {
                    if (pattern.StartsWith(RRULE))
                    {
                        var parts = pattern.Split(new char[] { ';', ':' });

                        var instance = 0;
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYDAY))
                            {
                                var days = part.Split(',');
                                foreach (var day in days)
                                {
                                    var dayValue = day.Substring(day.IndexOf("=") + 1);
                                    if (dayValue.StartsWith("1"))
                                    {
                                        instance = 1;
                                    }
                                    else if (dayValue.StartsWith("2"))
                                    {
                                        instance = 2;
                                    }
                                    else if (dayValue.StartsWith("3"))
                                    {
                                        instance = 3;
                                    }
                                    else if (dayValue.StartsWith("4"))
                                    {
                                        instance = 4;
                                    }

                                    break;
                                }
                                break;
                            }
                        }

                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYSETPOS))
                            {
                                var pos = part.Substring(part.IndexOf("=") + 1);

                                if (pos.Trim() == "-1")
                                {
                                    instance = 5;
                                }
                                else
                                {
                                    throw new NotSupportedException("Only 'BYSETPOS=-1' is allowed by Outlook, but it was: " + part);
                                }
                                break;
                            }
                        }

                        foreach (var part in parts)
                        {
                            if (part.StartsWith(FREQ))
                            {
                                switch (part.Substring(part.IndexOf('=') + 1))
                                {
                                    case DAILY:
                                        if (or.RecurrenceType != Outlook.OlRecurrenceType.olRecursDaily)
                                        {
                                            return false;
                                        }
                                        break;
                                    case WEEKLY:
                                        if (or.RecurrenceType != Outlook.OlRecurrenceType.olRecursWeekly)
                                        {
                                            return false;
                                        }
                                        break;
                                    case MONTHLY:
                                        if (instance == 0)
                                        {
                                            if (or.RecurrenceType != Outlook.OlRecurrenceType.olRecursMonthly)
                                            {
                                                return false;
                                            }
                                        }
                                        else
                                        {
                                            if (or.RecurrenceType != Outlook.OlRecurrenceType.olRecursMonthNth)
                                            {
                                                return false;
                                            }
                                            if (or.Instance != instance)
                                            {
                                                return false;
                                            }
                                        }
                                        break;
                                    case YEARLY:
                                        if (instance == 0)
                                        {
                                            if (or.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearly)
                                            {
                                                return false;
                                            }
                                        }
                                        else
                                        {
                                            if (or.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearNth)
                                            {
                                                return false;
                                            }
                                            if (or.Instance != instance)
                                            {
                                                return false;
                                            }
                                        }
                                        break;
                                    default: throw new NotSupportedException("RecurrenceType not supported by Outlook: " + part);

                                }
                                break;
                            }
                        }

                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYDAY))
                            {
                                Outlook.OlDaysOfWeek dayOfWeek = 0;
                                var days = part.Split(',');
                                foreach (var day in days)
                                {
                                    var dayValue = day.Substring(day.IndexOf("=") + 1);

                                    switch (dayValue.Trim(new char[] { '1', '2', '3', '4', ' ' }))
                                    {
                                        case MO: dayOfWeek |= Outlook.OlDaysOfWeek.olMonday; break;
                                        case TU: dayOfWeek |= Outlook.OlDaysOfWeek.olTuesday; break;
                                        case WE: dayOfWeek |= Outlook.OlDaysOfWeek.olWednesday; break;
                                        case TH: dayOfWeek |= Outlook.OlDaysOfWeek.olThursday; break;
                                        case FR: dayOfWeek |= Outlook.OlDaysOfWeek.olFriday; break;
                                        case SA: dayOfWeek |= Outlook.OlDaysOfWeek.olSaturday; break;
                                        case SU: dayOfWeek |= Outlook.OlDaysOfWeek.olSunday; break;

                                    }
                                    //Don't break because multiple days possible;
                                }

                                if (or.DayOfWeekMask != dayOfWeek && dayOfWeek != 0)
                                {
                                    if (or.DayOfWeekMask != dayOfWeek)
                                    {
                                        return false;
                                    }
                                }

                                break;
                            }
                        }

                        foreach (var part in parts)
                        {
                            if (part.StartsWith(INTERVAL))
                            {
                                var interval = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                if (or.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearly ||
                                    or.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearNth)
                                {
                                    interval *= 12; // must be expressed in months
                                }
                                if (or.Interval != interval)
                                {
                                    return false;
                                }
                                break;
                            }
                        }

                        foreach (var part in parts)
                        {
                            if (part.StartsWith(COUNT))
                            {
                                if (or.Occurrences != int.Parse(part.Substring(part.IndexOf('=') + 1)))
                                {
                                    return false;
                                }
                                break;
                            }
                            else if (part.StartsWith(UNTIL))
                            {
                                if (!or.NoEndDate)
                                {
                                    var d = or.PatternEndDate.Date.AddDays(oa.AllDayEvent ? 0 : 1);
                                    if (d != GetDateTime(part.Substring(part.IndexOf('=') + 1)))
                                    {
                                        return false;
                                    }
                                }
                                break;
                            }
                        }

                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYMONTHDAY))
                            {
                                if (or.DayOfMonth != int.Parse(part.Substring(part.IndexOf('=') + 1)))
                                {
                                    return false;
                                }
                                break;
                            }
                        }

                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYMONTH + "="))
                            {
                                if (or.MonthOfYear != int.Parse(part.Substring(part.IndexOf('=') + 1)))
                                {
                                    return false;
                                }
                                break;
                            }
                        }
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                oa.ToDebugLog();
                ga.ToDebugLog();
                Log.Error($"Error comparing recurrence: {ga.ToLogString()}: {ex.Message}");
            }
            finally
            {
                if (or != null)
                {
                    Marshal.ReleaseComObject(or);
                }
            }

            return true;
        }

        /// <summary>
        /// Update Recurrence pattern from Google by parsing the string, see also specification http://tools.ietf.org/html/rfc2445
        /// </summary>
        /// <param name="master"></param>
        /// <param name="slave"></param>
        public static bool UpdateRecurrence(Event master, Outlook.AppointmentItem slave)
        {
            if (IsSameRecurrence(master, slave))
            {
                return true;
            }

            var masterRecurrence = master.Recurrence;
            if ((masterRecurrence == null) || (masterRecurrence.Count < 1))
            {
                if (slave.IsRecurring && slave.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster)
                {
                    slave.ClearRecurrencePattern();                    
                }
                return true;
            }

            Outlook.RecurrencePattern rp = null;

            try
            {
                //OK, this call will convert non-recurring Outlook appointment to recurring
                rp = slave.GetRecurrencePattern();

                foreach (var pattern in master.Recurrence)
                {
                    if (pattern.StartsWith(RRULE))
                    {
                        var parts = pattern.Split(new char[] { ';', ':' });

                        #region BYDAY
                        var instance = 0;
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYDAY))
                            {
                                var days = part.Split(',');
                                foreach (var day in days)
                                {
                                    var dayValue = day.Substring(day.IndexOf("=") + 1);
                                    if (dayValue.StartsWith("1"))
                                    {
                                        instance = 1;
                                    }
                                    else if (dayValue.StartsWith("2"))
                                    {
                                        instance = 2;
                                    }
                                    else if (dayValue.StartsWith("3"))
                                    {
                                        instance = 3;
                                    }
                                    else if (dayValue.StartsWith("4"))
                                    {
                                        instance = 4;
                                    }

                                    break;
                                }
                                break;
                            }
                        }
                        #endregion

                        #region BYSETPOS
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYSETPOS))
                            {
                                var pos = part.Substring(part.IndexOf("=") + 1);

                                if (pos.Trim() == "-1")
                                {
                                    instance = 5;
                                }
                                else
                                {
                                    Log.Warning($"Synchronizing Google appointment {master.ToLogString()}, such recurrence is not possible at Outlook. Only 'BYSETPOS=-1' is allowed by Outlook, but it was: {pos} (see full parsed string: {part})");
                                    slave.ToDebugLog();
                                    master.ToDebugLog();
                                    return false;
                                }
                                break;
                            }
                        }
                        #endregion

                        #region FREQ
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(FREQ))
                            {
                                switch (part.Substring(part.IndexOf('=') + 1))
                                {
                                    case DAILY:
                                        rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursDaily;
                                        break;
                                    case WEEKLY:
                                        rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
                                        break;
                                    case MONTHLY:
                                        if (instance == 0)
                                        {
                                            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthly;
                                        }
                                        else
                                        {
                                            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthNth;
                                            rp.Instance = instance;
                                        }
                                        break;
                                    case YEARLY:
                                        if (instance == 0)
                                        {
                                            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearly;
                                        }
                                        else
                                        {
                                            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearNth;
                                            rp.Instance = instance;
                                        }
                                        break;
                                    default:
                                        Log.Warning($"Synchronizing Google appointment {master.ToLogString()}, such recurrence is not possible at Outlook. RecurrenceType not supported by Outlook (see full parsed string: {part})");
                                        slave.ToDebugLog();
                                        master.ToDebugLog();
                                        return false;
                                }

                                //In Microsoft help:
                                //You must set the RecurrenceType property before you set other properties 
                                //for a RecurrencePattern object.
                                if (master.Start != null && !string.IsNullOrEmpty(master.Start.Date))
                                {
                                    rp.StartTime = DateTime.Parse(master.Start.Date);
                                    rp.PatternStartDate = DateTime.Parse(master.Start.Date);
                                }
                                else if (master.Start != null && master.Start.DateTimeDateTimeOffset.HasValue)
                                {
                                    rp.StartTime = master.Start.DateTimeDateTimeOffset.Value.DateTime;
                                    rp.PatternStartDate = master.Start.DateTimeDateTimeOffset.Value.DateTime;
                                }

                                if (master.End != null && !string.IsNullOrEmpty(master.End.Date))
                                {
                                    rp.EndTime = DateTime.Parse(master.End.Date);
                                }

                                if (master.End != null && master.End.DateTimeDateTimeOffset.HasValue)
                                {
                                    rp.EndTime = master.End.DateTimeDateTimeOffset.Value.DateTime;
                                }

                                break;
                            }
                        }
                        #endregion

                        #region BYDAY
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYDAY))
                            {
                                Outlook.OlDaysOfWeek dayOfWeek = 0;
                                var days = part.Split(',');
                                foreach (var day in days)
                                {
                                    var dayValue = day.Substring(day.IndexOf("=") + 1);

                                    switch (dayValue.Trim(new char[] { '1', '2', '3', '4', ' ' }))
                                    {
                                        case MO: dayOfWeek |= Outlook.OlDaysOfWeek.olMonday; break;
                                        case TU: dayOfWeek |= Outlook.OlDaysOfWeek.olTuesday; break;
                                        case WE: dayOfWeek |= Outlook.OlDaysOfWeek.olWednesday; break;
                                        case TH: dayOfWeek |= Outlook.OlDaysOfWeek.olThursday; break;
                                        case FR: dayOfWeek |= Outlook.OlDaysOfWeek.olFriday; break;
                                        case SA: dayOfWeek |= Outlook.OlDaysOfWeek.olSaturday; break;
                                        case SU: dayOfWeek |= Outlook.OlDaysOfWeek.olSunday; break;
                                    }
                                    //Don't break because multiple days possible;
                                }

                                if (rp.DayOfWeekMask != dayOfWeek && dayOfWeek != 0)
                                {
                                    try
                                    {
                                        rp.DayOfWeekMask = dayOfWeek;
                                    }
                                    catch
                                    {
                                        Log.Warning($"Day of week mask {dayOfWeek} is invalid for {master.ToLogString()} (see full parsed string: {part})");
                                        slave.ToDebugLog();
                                        master.ToDebugLog();
                                        return false;
                                    }
                                }

                                break;
                            }
                        }
                        #endregion

                        #region INTERVAL
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(INTERVAL))
                            {
                                var interval = int.Parse(part.Substring(part.IndexOf('=') + 1));

                                if (rp.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearly ||
                                    rp.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearNth)
                                {
                                    if (interval > 8)
                                    {
                                        Log.Warning($"Synchronizing Google appointment {master.ToLogString()}, such recurrence is not possible at Outlook. Yearly recurrence and interval: {interval} (see full parsed string: {part})");
                                        slave.ToDebugLog();
                                        master.ToDebugLog();
                                        return false;
                                    }
                                    interval *= 12; // must be expressed in months
                                }
                                
                                try
                                {
                                    rp.Interval = interval;
                                }
                                catch (Exception ex)
                                {
                                    Log.Debug(ex, $"Error updating recurrence Interval {interval} for appointment {slave.ToLogString()}");
                                }
                                break;
                            }
                        }
                        #endregion

                        #region COUNT UNTIL
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(COUNT))
                            {
                                string parsedString = part.Substring(part.IndexOf('=') + 1);
                                try
                                {
                                    rp.Occurrences = int.Parse(parsedString);
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    Log.Warning($"Error parsing Recurrence {COUNT} value for Google Appointment {master.ToLogString()}, expected integer, but was: {parsedString} (see full parsed string: {part})\nException:{ex.Message}");
                                    Log.Debug(ex, "Exception");
                                    slave.ToDebugLog();
                                    master.ToDebugLog();
                                    return false;
                                }
                            }
                            else if (part.StartsWith(UNTIL))
                            {
                                //either UNTIL or COUNT may appear in a 'recur',
                                //but UNTIL and COUNT MUST NOT occur in the same 'recur'

                                DateTime dt = DateTime.MaxValue;
                                string parsedString = part.Substring(part.IndexOf('=') + 1);
                                try
                                {
                                    dt = GetDateTime(parsedString);
                                }
                                catch (Exception ex)
                                {
                                    Log.Warning($"Error parsing Recurrence {UNTIL} value for Google Appointment {master.ToLogString()}, expected DateTime, but was: {parsedString} (see full parsed string: {part})\nException: {ex.Message}");
                                    Log.Debug(ex, "Exception");
                                    slave.ToDebugLog();
                                    master.ToDebugLog();
                                    return false;
                                }

                                if (dt.Year < 4500)
                                {
                                    rp.PatternEndDate = dt;
                                }
                                else
                                {
                                    Log.Warning($"Synchronizing Google appointment {master.ToLogString()}, recurrence ends {dt}. Such value is not accepted by Outlook. You can change Google appointment and change recurrence it never ends or change end date of recurrence.");
                                    slave.ToDebugLog();
                                    master.ToDebugLog();
                                    return false;
                                }
                                break;
                            }
                        }
                        #endregion

                        #region BYMONTHDAY
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYMONTHDAY))
                            {
                                int dayOfMonth;
                                string parsedString = part.Substring(part.IndexOf('=') + 1);
                                try
                                {
                                    dayOfMonth = int.Parse(parsedString);
                                }
                                catch (Exception ex)
                                {
                                    Log.Warning($"Error parsing Recurrence {BYMONTHDAY} value for Google Appointment {master.ToLogString()}, expected integer, but was: {parsedString} (see full parsed string: {part})\nException:{ex.Message}");
                                    Log.Debug(ex, "Exception");
                                    slave.ToDebugLog();
                                    master.ToDebugLog();
                                    return false;
                                }


                                try
                                {                                    
                                    rp.DayOfMonth = dayOfMonth;
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    Log.Warning($"Day of month {dayOfMonth} is invalid for {master.ToLogString()}");
                                    Log.Debug(ex, "Exception");
                                    slave.ToDebugLog();
                                    master.ToDebugLog();
                                    return false;
                                }
                            }
                        }
                        #endregion

                        #region BYMONTH
                        foreach (var part in parts)
                        {
                            if (part.StartsWith(BYMONTH + "="))
                            {
                                rp.MonthOfYear = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }
                        #endregion
                        break;
                    }
                }
            }
            finally
            {
                if (rp != null)
                {
                    Marshal.ReleaseComObject(rp);
                }
            }

            return true;
        }

        internal static bool InSyncPeriod(DateTime d)
        {
            return InSyncPeriod(d, d);
        }

        private static bool InSyncPeriod(DateTime s, DateTime e)
        {
            if ((!Synchronizer.RestrictMonthsInPast || e >= DateTime.Now.AddMonths(-Synchronizer.MonthsInPast)) &&
                (!Synchronizer.RestrictMonthsInFuture || s <= DateTime.Now.AddMonths(Synchronizer.MonthsInFuture)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool UpdateException(Outlook.Exception master, Event slave, Synchronizer sync)
        {
            Outlook.AppointmentItem oa = null;

            try
            {
                oa = master.AppointmentItem;
            }
            catch
            {
            }

            try
            {
                if (oa != null)
                {
                    slave = sync.GetGoogleAppointment(slave.Id);
                    UpdateAppointment(oa, slave, true);
                    var updatedSlave = sync.SaveGoogleAppointment(slave);

                    var ret = updatedSlave != null && (updatedSlave != slave || updatedSlave.ETag != slave.ETag);
                    return ret;
                }
                else
                {
                    Log.Information($"Unable to find changed Outlook appointment: {slave.ToLogString()}");
                    slave.ToDebugLog();
                    return false;
                }
            }
            catch (COMException ex)
            {
                Log.Debug(ex, "Exception");
                slave.ToDebugLog();
                master.ToDebugLog();
                Log.Information($"Unable to find changed Outlook appointment: {slave.ToLogString()}");
                return false;
            }
            finally
            {
                if (oa != null)
                {
                    Marshal.ReleaseComObject(oa);
                }
            }
        }

        private static bool UpdateRecurrenceNotDeletedException(Outlook.Exception e, Outlook.AppointmentItem master, Event slave, ref List<Event> GoogleAppointmentExceptions, Synchronizer sync)
        {
            if (e.OriginalDate == null)
            {
                return false;
            }

            if (!InSyncPeriod(e.OriginalDate))
            {
                return false;
            }

            if (GoogleAppointmentExceptions != null)
            {
                for (var i = GoogleAppointmentExceptions.Count - 1; i >= 0; i--)
                {
                    var ga = GoogleAppointmentExceptions[i];

                    if (EqualDates(ga.OriginalStartTime, e.OriginalDate))
                    {
                        GoogleAppointmentExceptions.RemoveAt(i);
                        return UpdateException(e, ga, sync);
                    }
                }
            }

            var r = sync.GetGoogleAppointmentInstances(slave.Id);

            if (slave.Start != null)
            {
                if (slave.Start.DateTimeDateTimeOffset.HasValue)
                {
                    var s = slave.Start.DateTimeDateTimeOffset.Value.DateTime;
                    var se = new DateTime(
                                e.OriginalDate.Year,
                                e.OriginalDate.Month,
                                e.OriginalDate.Day,
                                s.Hour,
                                s.Minute,
                                s.Second,
                                s.Kind);
                    r.OriginalStart = se.ToString("yyyy-MM-ddTHH:mm:sszzz");
                }
                else if (slave.Start.Date != null)
                {
                    r.OriginalStart = e.OriginalDate.ToString("yyyy-MM-dd");
                }
            }
            var instances = r.Execute();

            var found = false;
            var changed = false;
            for (var i = instances.Items.Count - 1; i >= 0; i--)
            {
                var ga = instances.Items[i];

                if (EqualDates(ga.OriginalStartTime, e.OriginalDate))
                {
                    found = true;
                    changed |= UpdateException(e, ga, sync);
                }
            }
            if (found)
            {
                return changed;
            }

            Log.Information($"Unable to find Google event instance for Outlook appointment exception: {master.Subject} - {e.OriginalDate}");
            master.ToDebugLog();
            slave.ToDebugLog();

            return false;
        }

        private static bool EqualDates(EventDateTime g, DateTime o)
        {
            if (g == null)
            {
                return false;
            }

            if (o == null)
            {
                return false;
            }

            if (g.DateTimeDateTimeOffset != null)
            {
                if (o.Date == g.DateTimeDateTimeOffset.Value.Date)
                {
                    return true;
                }
            }
            else if (g.Date != null)
            {
                if (o.Date == DateTime.Parse(g.Date))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool UpdateRecurrenceDeletedException(Outlook.Exception oe, Outlook.AppointmentItem master, Event slave, ref List<Event> GoogleAppointmentExceptions, Synchronizer sync)
        {
            if (oe.OriginalDate == null)
            {
                Log.Debug($"Deleted Outlook Appointment Recurrence Exception doesn't have OriginalDate {master.Subject} ==> Skipping to Delete Exception");
                return false;
            }

            if (!InSyncPeriod(oe.OriginalDate))
            {
                Log.Debug($"Deleted Outlook Appointment Recurrence Exception not in sync period {master.Subject} - {oe.OriginalDate} ==> Skipping to Delete");
                return false;
            }

            if (GoogleAppointmentExceptions != null)
            {
                for (var i = GoogleAppointmentExceptions.Count - 1; i >= 0; i--)
                {
                    var ge1 = GoogleAppointmentExceptions[i];
                    var ge = sync.GetGoogleAppointment(ge1.Id);

                    if (EqualDates(ge.OriginalStartTime, oe.OriginalDate))
                    {
                        if (ge.Status != null && ge.Status == "cancelled")
                        {
                            //deleted google exception already exists do not create new one
                            GoogleAppointmentExceptions.RemoveAt(i);
                            return false;
                        }
                        else
                        {
                            //google exception already exists, delete it
                            ge.Status = "cancelled";
                            sync.SaveGoogleAppointment(ge);
                            GoogleAppointmentExceptions.RemoveAt(i);
                            Log.Information($"Deleted recurrence exception from Google: {ge1.ToLogString()}");
                            return true;
                        }
                    }
                }
            }

            var instancesRequest = sync.GetGoogleAppointmentInstances(slave.Id);

            if (slave.Start != null)
            {
                if (slave.Start.DateTimeDateTimeOffset.HasValue)
                {
                    var s = slave.Start.DateTimeDateTimeOffset.Value.DateTime;
                    var se = new DateTime(
                                oe.OriginalDate.Year,
                                oe.OriginalDate.Month,
                                oe.OriginalDate.Day,
                                s.Hour,
                                s.Minute,
                                s.Second,
                                s.Kind);
                    instancesRequest.OriginalStart = se.ToString("yyyy-MM-ddTHH:mm:sszzz");
                }
                else if (slave.Start.Date != null)
                {
                    instancesRequest.OriginalStart = oe.OriginalDate.ToString("yyyy-MM-dd");
                }
            }
            var instances = instancesRequest.Execute();

            var found = false;
            for (var i = instances.Items.Count - 1; i >= 0; i--)
            {
                var ge1 = instances.Items[i];
                var ge = sync.GetGoogleAppointment(ge1.Id);

                if (EqualDates(ge.OriginalStartTime, oe.OriginalDate))
                {
                    if (ge.Status == null || ge.Status != "cancelled")
                    {
                        //google instance already exists, delete it
                        ge.Status = "cancelled";
                        sync.SaveGoogleAppointment(ge);
                        Log.Information($"Deleted recurrence exception from Google: {ge1.ToLogString()}");
                        found = true;
                    }
                }
            }
            if (found)
            {
                return true;
            }


            Log.Information($"Unable to find Google event instance for deleted Outlook appointment exception: {master.Subject} - {oe.OriginalDate}");
            //Log.Debug(master);
            if (!string.IsNullOrEmpty(instancesRequest.OriginalStart))
            {
                Log.Debug($"Original start date: {instancesRequest.OriginalStart}");
            }
            slave.ToDebugLog();

            return false;
        }

        internal static bool UpdateRecurrenceExceptions(Outlook.AppointmentItem master, Event slave, ref List<Event> GoogleAppointmentExceptions, Synchronizer sync)
        {
            Outlook.RecurrencePattern rp = null;
            Outlook.Exceptions exceptions = null;
            var ret = false;

            try
            {
                rp = master.GetRecurrence();
                exceptions = rp.Exceptions;
                if (exceptions == null || exceptions.Count == 0)
                    ret = true;
                else
                {
                    for (var i = exceptions.Count; i > 0; i--)
                    {
                        Outlook.Exception oe = null;

                        try
                        {
                            oe = exceptions[i];
                            if (oe.Deleted)
                            {
                                if (UpdateRecurrenceDeletedException(oe, master, slave, ref GoogleAppointmentExceptions, sync))
                                {
                                    ret = true;
                                }
                            }
                            else
                            {
                                if (UpdateRecurrenceNotDeletedException(oe, master, slave, ref GoogleAppointmentExceptions, sync))
                                {
                                    ret = true;
                                }
                            }
                        }
                        finally
                        {
                            if (oe != null)
                            {
                                Marshal.ReleaseComObject(oe);
                            }
                        }
                    }
                }


                if (GoogleAppointmentExceptions != null)
                {
                    //after sync, some Google exceptions left, they need to be deleted
                    for (var i = GoogleAppointmentExceptions.Count - 1; i >= 0; i--)
                    {
                        var ga = GoogleAppointmentExceptions[i];

                        var ga1 = sync.GetGoogleAppointment(ga.Id);

                        //undelete deleted Google exception
                        DateTime date = AppointmentSync.outlookDateInvalid;//set any invalid date
                        if (ga1.OriginalStartTime != null)
                        {
                            var oa = AppointmentPropertiesUtils.GetOccurrence(ga1.OriginalStartTime, rp, ref date, master);
                            if (ga1.Status == "cancelled" && oa != null && oa.MeetingStatus != Outlook.OlMeetingStatus.olMeetingCanceled && oa.MeetingStatus != Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
                            {
                                ga1.Status = slave.Status;
                                UpdateAppointment(oa, ga1, true);
                                try
                                {
                                    sync.SaveGoogleAppointment(ga1);
                                }
                                catch (Exception ex)
                                {
                                    Log.Debug(ex, "Exception");
                                    ga1.ToDebugLog();
                                    master.ToDebugLog();
                                    slave.ToDebugLog();
                                }
                                GoogleAppointmentExceptions.RemoveAt(i);
                                ret = true;
                            }
                        }
                    }
                }
            }
            finally
            {
                if (exceptions != null)
                {
                    Marshal.ReleaseComObject(exceptions);
                }

                if (rp != null)
                {
                    Marshal.ReleaseComObject(rp);
                }
            }
            return ret;
        }

        private static bool UpdateRecurrenceException(Event master, ref Outlook.AppointmentItem slave, Synchronizer sync)
        {
            var ret = false;

            Outlook.AppointmentItem oa = null;
            Outlook.RecurrencePattern rp = null;
            try
            {                
                rp = slave.GetRecurrence();
                var date = AppointmentSync.outlookDateInvalid;//set any invalid date
                if (master.OriginalStartTime != null)
                {
                    try
                    {
                        oa = AppointmentPropertiesUtils.GetOccurrence(master.OriginalStartTime, rp, ref date, slave);
                    }
                    catch (COMException ex) when ((uint)ex.HResult == 0x80004005)
                    {
                        Log.Debug($"Google Appointment with OriginalEvent found, but Outlook occurrence not found: {master.Summary} - {master.OriginalStartTime.DateTimeDateTimeOffset}: {ex}");
                    }

                }
               
                

                var reloaded_master = sync.GetGoogleAppointment(master.Id);

                if (reloaded_master != null)
                {
                    if (oa != null)
                    {
                        if (reloaded_master.Status.Equals("cancelled"))
                        {
                            var txt = oa.ToLogString();
                            oa.Delete();
                            slave.Save();

                            Log.Information($"Deleted one recurrence from Outlook appointment: {txt}");
                            ret = true;
                        }
                        else
                        {
                            if (sync.UpdateAppointment(ref reloaded_master, ref oa, null))
                            {
                                oa.Save();
                                Log.Information($"Updated recurrence exception from Google to Outlook: {reloaded_master.ToLogString()}");
                                ret = true;
                            }
                        }
                        //ret = true;
                    }
                    else
                    {
                        Log.Debug($"Google Appointment with OriginalEvent found, but Outlook occurrence not found: {reloaded_master.Summary} - {reloaded_master.OriginalStartTime.DateTimeDateTimeOffset}");
                    }
                }
                else
                {
                    Log.Warning($"Error updating recurrence exception from Google to Outlook (couldn't be reload from Google): {master.ToLogString()}");
                }
            }
            finally
            {
                if (oa != null)
                {
                    Marshal.ReleaseComObject(oa);
                }

                if (rp != null)
                {
                    Marshal.ReleaseComObject(rp);
                }
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(slave);

            if (slave != null)
            {
                Marshal.ReleaseComObject(slave);
            }

            slave = Synchronizer.OutlookNameSpace.GetItemFromID(oid);

            return ret;
        }

        private static bool CheckIfOutlookExceptionsNeedsRebuild(Outlook.AppointmentItem slave)
        {
            Outlook.RecurrencePattern rp = null;
            Outlook.Exceptions e = null;
            try
            {
                rp = slave.GetRecurrence();
                e = rp.Exceptions;

                if (e.Count == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            finally
            {
                if (e != null)
                {
                    Marshal.ReleaseComObject(e);
                }

                if (rp != null)
                {
                    Marshal.ReleaseComObject(rp);
                }
            }
        }

        private static void ClearOutlookExceptions(ref Outlook.AppointmentItem oa)
        {            
            Outlook.RecurrencePattern rp = null;
            try
            {
                rp = oa.GetRecurrence();
                if (rp != null)
                {
                    var currentPatternStartDate = rp.PatternStartDate;
                    rp.PatternStartDate = currentPatternStartDate.AddYears(-1);
                    rp.PatternStartDate = currentPatternStartDate;
                }
                //oa.ClearRecurrencePattern();//ToDo: Maybe additionally set IsRecurring flag to false? but then the recurrence must be again setup later
                oa.Save();
            }
            finally
            {
                if (rp != null)
                {
                    Marshal.ReleaseComObject(rp);
                }
            }
        }

        internal static bool UpdateRecurrenceExceptions(List<Event> googleRecurrenceExceptions, ref Outlook.AppointmentItem slave, Synchronizer sync)
        {
            var ret = false;

            if (CheckIfOutlookExceptionsNeedsRebuild(slave))
            {
                ClearOutlookExceptions(ref slave);                
                ret = true;
            }

            if (googleRecurrenceExceptions == null || googleRecurrenceExceptions.Count == 0)
            {
                ret = true;
            }
            else
            {
                for (var i = 0; i < googleRecurrenceExceptions.Count; i++)
                {
                    var ga = googleRecurrenceExceptions[i];
                    if (UpdateRecurrenceException(ga, ref slave, sync))
                    {
                        ret = true;
                    }
                }
            }

            return ret;
        }

        private static DateTime GetDateTime(string dateTime)
        {
            var format = dateFormat;
            if (dateTime.Contains("T"))
            {
                format += "'T'" + timeFormat;
            }

            if (dateTime.EndsWith("Z"))
            {
                format += "'Z'";
            }

            return DateTime.ParseExact(dateTime, format, new System.Globalization.CultureInfo("en-US"));
        }

        internal static bool IsOrganizer(string email)
        {
            if (email != null)
            {
                var userName = Synchronizer.UserName.Trim().ToLower().Replace("@googlemail.", "@gmail.");
                email = email.Trim().ToLower().Replace("@googlemail.", "@gmail.");
                return email.Equals(userName, StringComparison.InvariantCultureIgnoreCase);
            }
            return false;
        }
    }
}
