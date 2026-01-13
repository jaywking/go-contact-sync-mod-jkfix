using Google.Apis.Calendar.v3.Data;
using Serilog;

namespace GoContactSyncMod
{
    public static class EventExtensions
    {
        private static string GetTime(this Event ga)
        {
            var ret = string.Empty;

            if (ga.Start != null && !string.IsNullOrEmpty(ga.Start.Date))
            {
                ret += ga.Start.Date;
            }
            else if (ga.Start != null && ga.Start.DateTimeDateTimeOffset != null)
            {
                ret += ga.Start.DateTimeDateTimeOffset.Value.ToString();
            }

            if (ga.Recurrence != null && ga.Recurrence.Count > 0)
            {
                ret += " Recurrence"; //ToDo: Return Recurrence Start/End
            }

            return ret;
        }

        public static void ToDebugLog(this Event e)
        {
            Log.Debug("*** Google event ***");
            Log.Debug(" - AnyoneCanAddSelf: " + (e.AnyoneCanAddSelf != null ? e.AnyoneCanAddSelf.ToString() : "null"));
            if (e.Attachments != null)
            {
                Log.Debug(" - Attachments:");
                foreach (var a in e.Attachments)
                {
                    Log.Debug("  - Title: " + (a.Title ?? "null"));
                }
            }
            if (e.Attendees != null)
            {
                Log.Debug(" - Attendees:");
                foreach (var a in e.Attendees)
                {
                    Log.Debug("  - DisplayName: " + (a.DisplayName ?? "null"));
                }
            }
            Log.Debug(" - AttendeesOmitted: " + (e.AttendeesOmitted != null ? e.AttendeesOmitted.ToString() : "null"));
            Log.Debug(" - ColorId: " + (e.ColorId ?? "null"));
            Log.Debug(" - Created: " + (e.CreatedDateTimeOffset != null ? e.CreatedDateTimeOffset.ToString() : "null"));
            Log.Debug(" - CreatedRaw: " + (e.CreatedRaw ?? "null"));
            if (e.Creator != null)
            {
                Log.Debug(" - Creator:");
                Log.Debug("  - DisplayName: " + (e.Creator.DisplayName ?? "null"));
            }
            Log.Debug($" - Description: {e.Description.RemoveNewLines()}");
            if (e.End != null)
            {
                Log.Debug(" - End:");
                if (!string.IsNullOrEmpty(e.End.Date))
                {
                    Log.Debug("  - Date: " + e.End.Date);
                }

                if (e.End.DateTimeDateTimeOffset != null)
                {
                    Log.Debug("  - DateTime: " + e.End.DateTimeDateTimeOffset.Value.ToString());
                }

                if (!string.IsNullOrEmpty(e.End.TimeZone))
                {
                    Log.Debug("  - TimeZone: " + e.End.TimeZone);
                }
            }
            Log.Debug(" - EndTimeUnspecified: " + (e.EndTimeUnspecified != null ? e.EndTimeUnspecified.ToString() : "null"));
            if (e.ExtendedProperties != null)
            {
                Log.Debug(" - ExtendedProperties:");
                if (e.ExtendedProperties.Shared != null)
                {
                    Log.Debug("  - Shared:");
                    foreach (var p in e.ExtendedProperties.Shared)
                    {
                        Log.Debug("   - Key: " + (p.Key ?? "null"));
                        Log.Debug("   - Value: " + (p.Value ?? "null"));
                    }
                }
                if (e.ExtendedProperties.Private__ != null)
                {
                    Log.Debug("  - Private__:");
                    foreach (var p in e.ExtendedProperties.Private__)
                    {
                        Log.Debug("   - Key: " + (p.Key ?? "null"));
                        Log.Debug("   - Value: " + (p.Value ?? "null"));
                    }
                }
            }
            if (e.Gadget != null)
            {
                Log.Debug(" - Gadget:");
                if (!string.IsNullOrEmpty(e.Gadget.Title))
                {
                    Log.Debug("  - Title: " + e.Gadget.Title);
                }
            }
            Log.Debug(" - GuestsCanInviteOthers: " + (e.GuestsCanInviteOthers != null ? e.GuestsCanInviteOthers.ToString() : "null"));
            Log.Debug(" - GuestsCanModify: " + (e.GuestsCanModify != null ? e.GuestsCanModify.ToString() : "null"));
            Log.Debug(" - GuestsCanSeeOtherGuests: " + (e.GuestsCanSeeOtherGuests != null ? e.GuestsCanSeeOtherGuests.ToString() : "null"));
            Log.Debug(" - HangoutLink: " + (e.HangoutLink ?? "null"));
            Log.Debug(" - HtmlLink: " + (e.HtmlLink ?? "null"));
            Log.Debug(" - ICalUID: " + (e.ICalUID ?? "null"));
            Log.Debug(" - Id: " + (e.Id ?? "null"));
            Log.Debug(" - Kind: " + (e.Kind ?? "null"));
            Log.Debug(" - Location: " + (e.Location ?? "null"));
            Log.Debug(" - Locked: " + (e.Locked != null ? e.Locked.ToString() : "null"));
            if (e.Organizer != null)
            {
                Log.Debug(" - Organizer:");
                Log.Debug("  - DisplayName: " + (e.Organizer.DisplayName ?? "null"));
            }
            if (e.OriginalStartTime != null)
            {
                Log.Debug(" - OriginalStartTime:");
                if (!string.IsNullOrEmpty(e.OriginalStartTime.Date))
                {
                    Log.Debug("  - Date: " + e.OriginalStartTime.Date);
                }

                if (e.OriginalStartTime.DateTimeDateTimeOffset != null)
                {
                    Log.Debug("  - DateTime: " + e.OriginalStartTime.DateTimeDateTimeOffset.Value.ToString());
                }

                if (!string.IsNullOrEmpty(e.OriginalStartTime.TimeZone))
                {
                    Log.Debug("  - TimeZone: " + e.OriginalStartTime.TimeZone);
                }
            }
            Log.Debug(" - PrivateCopy: " + (e.PrivateCopy != null ? e.PrivateCopy.ToString() : "null"));
            if (e.Recurrence != null)
            {
                Log.Debug(" - Recurrence:");
                foreach (var r in e.Recurrence)
                {
                    Log.Debug("  - : " + r);
                }
            }
            Log.Debug(" - RecurringEventId: " + (e.RecurringEventId ?? "null"));
            if (e.Reminders != null)
            {
                Log.Debug(" - Reminders:");
                if (e.Reminders.UseDefault != null)
                {
                    Log.Debug("  - UseDefault: " + e.Reminders.UseDefault.ToString());
                }

                if (e.Reminders.Overrides != null)
                {
                    Log.Debug("  - Overrides:");
                    foreach (var o in e.Reminders.Overrides)
                    {
                        Log.Debug("   - Minutes: " + (o.Minutes != null ? o.Minutes.ToString() : "null"));
                    }
                }
            }
            Log.Debug(" - Sequence: " + (e.Sequence != null ? e.Sequence.ToString() : "null"));
            if (e.Source != null)
            {
                Log.Debug(" - Source:");
                Log.Debug("  - Url: " + (e.Source.Url ?? "null"));
            }
            if (e.Start != null)
            {
                Log.Debug(" - Start:");
                if (!string.IsNullOrEmpty(e.Start.Date))
                {
                    Log.Debug("  - Date: " + e.Start.Date);
                }

                if (e.Start.DateTimeDateTimeOffset != null)
                {
                    Log.Debug("  - DateTime: " + e.Start.DateTimeDateTimeOffset.Value.ToString());
                }

                if (!string.IsNullOrEmpty(e.Start.TimeZone))
                {
                    Log.Debug("  - TimeZone: " + e.Start.TimeZone);
                }
            }
            Log.Debug(" - Status: " + (e.Status ?? "null"));
            Log.Debug($" - Summary: {e.Summary.RemoveNewLines()}");
            Log.Debug(" - Transparency: " + (e.Transparency ?? "null"));
            Log.Debug(" - Updated: " + (e.UpdatedDateTimeOffset != null ? e.UpdatedDateTimeOffset.ToString() : "null"));
            Log.Debug(" - UpdatedRaw: " + (e.UpdatedRaw ?? "null"));
            Log.Debug(" - Visibility: " + (e.Visibility ?? "null"));
            Log.Debug("*** Google event ***");
        }

        public static string ToLogString(this Event ga)
        {
            var summary = ga.Summary;
            var time = ga.GetTime();

            if (string.IsNullOrWhiteSpace(summary))
            {
                if (string.IsNullOrWhiteSpace(time))
                {
                    if (ga.Status.Equals("cancelled"))
                    {
                        if (ga.OriginalStartTime != null)
                        {
                            if (!string.IsNullOrEmpty(ga.OriginalStartTime.Date))
                            {
                                return $"Cancelled - {ga.OriginalStartTime.Date}";
                            }
                            else if (ga.OriginalStartTime.DateTimeDateTimeOffset != null)
                            {
                                return $"Cancelled - {ga.OriginalStartTime.DateTimeDateTimeOffset.Value}";
                            }
                            else
                            {
                                return string.Empty;
                            }
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
                else
                {
                    return time;
                }
            }
            else
            {
                if (string.IsNullOrWhiteSpace(time))
                {
                    return summary;
                }
                else
                {
                    return $"{summary} - {time}";
                }
            }
        }
    }
}

