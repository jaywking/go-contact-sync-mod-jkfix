using Serilog;
using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    public static class AppointmentItemExtensions
    {
        public static string ToLogString(this Outlook.AppointmentItem oa)
        {
            if (oa == null)
            {
                return string.Empty;
            }

            string subject;
            try
            {
                subject = oa.Subject;
                if (subject != null)
                {
                    subject = subject.RemoveNewLines();
                }
                else
                {
                    subject = string.Empty;
                }
            }
            catch
            {
                subject = string.Empty;
            }

            string time;
            try
            {
                var st = oa.Start;
                if (st != null)
                {
                    if (oa.AllDayEvent)
                    {
                        time = st.ToString("d");
                    }
                    else
                    {
                        time = st.ToString("g");
                    }
                }
                else
                {
                    time = string.Empty;
                }
            }
            catch
            {
                time = string.Empty;
            }


            string recurrence;
            try
            {
                recurrence = oa.IsRecurring && oa.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster ? " Recurrence":string.Empty;
            }
            catch
            {
                recurrence = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(subject))
            {
                return time + recurrence;
            }

            if (string.IsNullOrWhiteSpace(time))
            {
                return subject + recurrence;
            }

            return $"{subject} - {time}" + recurrence;
        }

        public static void ToDebugLog(this Outlook.AppointmentItem e, bool printRecurrenceExceptions = true)
        {
            try
            {
                Log.Debug("*** Outlook appointment ***");

                Log.Debug(" - AllDayEvent: " + e.AllDayEvent ?? "null");

                Outlook.Attachments attachments = null;
                try
                {
                    attachments = e.Attachments;
                    if (attachments != null)
                    {
                        if (attachments.Count > 0)
                        {
                            Log.Debug($" - Attachments.Count: {attachments.Count}");
                            for (var i = attachments.Count; i > 0; i--)
                            {
                                Outlook.Attachment a = null;
                                try
                                {
                                    a = attachments[i];
                                    Log.Debug("  - Attachment: " + (a.DisplayName ?? "null"));
                                }
                                finally
                                {
                                    if (a != null)
                                    {
                                        Marshal.ReleaseComObject(a);
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    //Do Nothing
                }
                finally
                {
                    if (attachments != null)
                    {
                        Marshal.ReleaseComObject(attachments);
                    }
                }

                Log.Debug(" - AutoResolvedWinner: " + e.AutoResolvedWinner ?? "null");
                Log.Debug(" - BillingInformation: " + e.BillingInformation ?? "null");
                Log.Debug(" - Body: " + e.Body.Truncate(20) ?? "null");
                Log.Debug(" - BusyStatus: " + e.BusyStatus ?? "null");
                Log.Debug(" - Categories: " + e.Categories ?? "null");
                Log.Debug(" - Companies: " + e.Companies ?? "null");
                Log.Debug(" - ConferenceServerAllowExternal: " + e.ConferenceServerAllowExternal ?? "null");
                //do not access ConferenceServerPassword it is deprecated field
                //Log.Debug(" - ConferenceServerPassword: " + e.ConferenceServerPassword ?? "null");

                //Accessing ConversationID, ConversationIndex and ConversationTopic  for some appointments throws 
                //Member not found exception (0x80020003)
                try
                {
                    Log.Debug(" - ConversationID: " + e.ConversationID ?? "null");
                    Log.Debug(" - ConversationIndex: " + e.ConversationIndex ?? "null");
                    Log.Debug(" - ConversationTopic: " + e.ConversationTopic ?? "null");
                }
                catch
                {
                    //Do Nothing
                }

                Log.Debug(" - CreationTime: " + e.CreationTime ?? "null");
                Log.Debug(" - DownloadState: " + e.DownloadState ?? "null");
                Log.Debug(" - Duration: " + e.Duration ?? "null");
                Log.Debug(" - End: " + e.End ?? "null");

                // EndInEndTimeZone was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                try
                {
                    Log.Debug(" - EndInEndTimeZone: " + e.EndInEndTimeZone ?? "null");
                }
                catch (AccessViolationException)
                {
                    //Do Nothing
                }

                // EndTimeZone was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                Outlook.TimeZone etz = null;
                try
                {
                    etz = e.EndTimeZone;
                    if (etz != null)
                    {
                        Log.Debug(" - EndTimeZone: " + etz.ID ?? "null");
                    }
                }
                catch (AccessViolationException)
                {
                    //Do Nothing
                }
                finally
                {
                    if (etz != null)
                    {
                        Marshal.ReleaseComObject(etz);
                    }
                }

                // EndUTC was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                try
                {
                    Log.Debug(" - EndUTC: " + e.EndUTC ?? "null");
                }
                catch (AccessViolationException)
                {
                    //Do Nothing
                }

                // ForceUpdateToAllAttendees was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                try
                {
                    Log.Debug(" - ForceUpdateToAllAttendees: " + e.ForceUpdateToAllAttendees ?? "null");
                }
                catch (AccessViolationException)
                {
                    //Do Nothing
                }

                Log.Debug(" - Importance: " + e.Importance ?? "null");
                Log.Debug(" - InternetCodepage: " + e.InternetCodepage ?? "null");
                Log.Debug(" - IsConflict: " + e.IsConflict ?? "null");
                Log.Debug(" - IsOnlineMeeting: " + e.IsOnlineMeeting ?? "null");
                Log.Debug(" - IsRecurring: " + e.IsRecurring ?? "null");
                Log.Debug(" - LastModificationTime: " + e.LastModificationTime ?? "null");
                Log.Debug(" - Location: " + e.Location ?? "null");
                Log.Debug(" - MarkForDownload: " + e.MarkForDownload ?? "null");
                Log.Debug(" - MeetingStatus: " + e.MeetingStatus ?? "null");
                Log.Debug(" - MeetingWorkspaceURL: " + e.MeetingWorkspaceURL ?? "null");
                Log.Debug(" - Mileage: " + e.Mileage ?? "null");
                Log.Debug(" - NetMeetingDocPathName: " + e.NetMeetingDocPathName ?? "null");
                Log.Debug(" - NetMeetingOrganizerAlias: " + e.NetMeetingOrganizerAlias ?? "null");
                Log.Debug(" - NetMeetingServer: " + e.NetMeetingServer ?? "null");
                Log.Debug(" - NetMeetingType: " + e.NetMeetingType ?? "null");
                Log.Debug(" - NetShowURL: " + e.NetShowURL ?? "null");
                Log.Debug(" - NoAging: " + e.NoAging ?? "null");
                Log.Debug(" - OptionalAttendees: " + e.OptionalAttendees ?? "null");
                Log.Debug(" - Organizer: " + e.Organizer ?? "null");

                Outlook.Recipients recipients = null;
                try
                {
                    recipients = e.Recipients;
                    if (recipients != null)
                    {
                        if (recipients.Count > 0)
                        {
                            Log.Debug($" - Recipients.Count: {recipients.Count}");
                            for (var i = recipients.Count; i > 0; i--)
                            {
                                Outlook.Recipient a = null;
                                try
                                {
                                    a = recipients[i];
                                    Log.Debug("  - Recipients: " + (a.Name ?? "null"));
                                }
                                finally
                                {
                                    if (a != null)
                                    {
                                        Marshal.ReleaseComObject(a);
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    //Do Nothing
                }
                finally
                {
                    if (recipients != null)
                    {
                        Marshal.ReleaseComObject(recipients);
                    }
                }

                Log.Debug(" - RecurrenceState: " + e.RecurrenceState ?? "null");
                Log.Debug(" - ReminderMinutesBeforeStart: " + e.ReminderMinutesBeforeStart ?? "null");
                Log.Debug(" - ReminderOverrideDefault: " + e.ReminderOverrideDefault ?? "null");
                Log.Debug(" - ReminderPlaySound: " + e.ReminderPlaySound ?? "null");
                Log.Debug(" - ReminderSet: " + e.ReminderSet ?? "null");
                Log.Debug(" - ReminderSoundFile: " + e.ReminderSoundFile ?? "null");
                Log.Debug(" - ReplyTime: " + e.ReplyTime ?? "null");
                Log.Debug(" - RequiredAttendees: " + e.RequiredAttendees ?? "null");
                Log.Debug(" - Resources: " + e.Resources ?? "null");
                Log.Debug(" - ResponseRequested: " + e.ResponseRequested ?? "null");
                Log.Debug(" - ResponseStatus: " + e.ResponseStatus ?? "null");
                Log.Debug(" - Saved: " + e.Saved ?? "null");
                Log.Debug(" - Sensitivity: " + e.Sensitivity ?? "null");
                Log.Debug(" - Size: " + e.Size ?? "null");
                Log.Debug(" - Start: " + e.Start ?? "null");

                // StartInStartTimeZone was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                try
                {
                    Log.Debug(" - StartInStartTimeZone: " + e.StartInStartTimeZone ?? "null");
                }
                catch (AccessViolationException)
                {
                    //Do Nothing
                }

                // StartTimeZone was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                Outlook.TimeZone stz = null;
                try
                {
                    stz = e.StartTimeZone;
                    if (stz != null)
                    {
                        Log.Debug(" - StartTimeZone: " + stz.ID ?? "null");
                    }
                }
                catch (AccessViolationException)
                {
                    //Do Nothing
                }
                finally
                {
                    if (stz != null)
                    {
                        Marshal.ReleaseComObject(stz);
                    }
                }

                // StartUTC was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                try
                {
                    Log.Debug(" - StartUTC: " + e.StartUTC ?? "null");
                }
                catch (AccessViolationException)
                {
                    //Do Nothing
                }

                Log.Debug(" - Subject: " + e.Subject ?? "null");
                Log.Debug(" - UnRead: " + e.UnRead ?? "null");

                Outlook.UserProperties up = null;
                try
                {
                    up = e.UserProperties;
                    if (up != null)
                    {
                        if (up.Count > 0)
                        {
                            Log.Debug($" - UserProperties.Count: {up.Count}");
                            for (var i = up.Count; i > 0; i--)
                            {
                                Outlook.UserProperty a = null;
                                try
                                {
                                    a = up[i];
                                    string v = Convert.ToString(a.Value);
                                    Log.Debug("  - UserProperty: " + (a.Name ?? "null") + " (" + (v.Truncate(20) ?? "null") + ")");
                                }
                                finally
                                {
                                    if (a != null)
                                    {
                                        Marshal.ReleaseComObject(a);
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    //Do Nothing
                }
                finally
                {
                    if (up != null)
                    {
                        Marshal.ReleaseComObject(up);
                    }
                }

                if (printRecurrenceExceptions)
                {
                    if (e.IsRecurring)
                    {
                        Outlook.RecurrencePattern r = null;
                        try
                        {
                            r = e.GetRecurrence();
                            if (r != null)
                            {
                                Log.Debug("** Outlook appointment recurrence **");
                                Log.Debug("  - DayOfMonth: " + r.DayOfMonth ?? "null");
                                Log.Debug("  - DayOfWeekMask: " + r.DayOfWeekMask ?? "null");
                                Log.Debug("  - Duration: " + r.Duration ?? "null");
                                Log.Debug("  - EndTime: " + r.EndTime ?? "null");
                                Log.Debug("  - Instance: " + r.Instance ?? "null");
                                Log.Debug("  - Interval: " + r.Interval ?? "null");
                                Log.Debug("  - MonthOfYear: " + r.MonthOfYear ?? "null");
                                Log.Debug("  - NoEndDate: " + r.NoEndDate ?? "null");
                                Log.Debug("  - Occurrences: " + r.Occurrences ?? "null");
                                Log.Debug("  - PatternEndDate: " + r.PatternEndDate ?? "null");
                                Log.Debug("  - PatternStartDate: " + r.PatternStartDate ?? "null");
                                Log.Debug("  - RecurrenceType: " + r.RecurrenceType ?? "null");
                                Log.Debug("  - Regenerate: " + r.Regenerate ?? "null");
                                Log.Debug("  - StartTime: " + r.StartTime ?? "null");
                                Log.Debug("** Outlook appointment recurrence **");

                                Outlook.Exceptions exceptions = null;
                                try
                                {
                                    exceptions = r.Exceptions;
                                    if (exceptions != null)
                                    {
                                        if (exceptions.Count > 0)
                                        {
                                            Log.Debug($"  - Exceptions.Count: {exceptions.Count}");
                                            for (var i = exceptions.Count; i > 0; i--)
                                            {
                                                Outlook.Exception ex = null;

                                                try
                                                {
                                                    ex = exceptions[i];
                                                    ex.ToDebugLog();
                                                }
                                                catch (COMException)
                                                {
                                                    //Do nothing
                                                }
                                                finally
                                                {
                                                    if (ex != null)
                                                    {
                                                        Marshal.ReleaseComObject(ex);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception)
                                {
                                    //Do nothing
                                }
                                finally
                                {
                                    if (exceptions != null)
                                    {
                                        Marshal.ReleaseComObject(exceptions);
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            //Do nothing
                        }
                        finally
                        {
                            if (r != null)
                            {
                                Marshal.ReleaseComObject(r);
                            }
                        }
                    }
                }
                Log.Debug("*** Outlook appointment ***");
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception logging details of an Outlook appointment");
            }
        }

        public static Outlook.RecurrencePattern GetRecurrence(this Outlook.AppointmentItem oa)
        {
            //if (oa.IsRecurring)
            //{
                return oa.GetRecurrencePattern();
            //}

            //throw new ApplicationException("Get Recurrence");
        }

        /*public static bool IsDeleted(this Outlook.AppointmentItem oa)
        {
            Outlook.RecurrencePattern rp = null;
            Outlook.Exceptions exceptions = null;

            try
            {
                if (!oa.IsRecurring)
                {
                    return false;
                }

                rp = oa.GetRecurrence();
                if (rp == null)
                {
                    return false;
                }

                exceptions = rp.Exceptions;
                if (exceptions == null || exceptions.Count == 0)
                {
                    return false;
                }

                if (rp.Occurrences != exceptions.Count)
                {
                    return false;
                }

                for (var i = 1; i <= exceptions.Count; i++)
                {
                    Outlook.Exception exception = null;

                    try
                    {
                        exception = exceptions[i];

                        if (!exception.Deleted)
                        {
                            return false;
                        }
                    }
                    finally
                    {
                        if (exception != null)
                        {
                            Marshal.ReleaseComObject(exception);
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
            return true;
        }*/
    }
}
