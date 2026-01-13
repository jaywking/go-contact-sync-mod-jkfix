using Google.Apis.Calendar.v3.Data;
using Serilog;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal static class AppointmentPropertiesUtils
    {
        internal static string GetOutlookId(Outlook.AppointmentItem oa)
        {
            return oa.EntryID;
        }

        internal static string GetGoogleId(Event ga)
        {
            var id = ga.Id.ToString();
            if (id == null)
            {
                throw new Exception();
            }
            return id;
        }

        internal static bool SetGoogleOutlookId(Event ga, Outlook.AppointmentItem oa)
        {
            var oid = GetOutlookId(oa);
            if (oid == null)
            {
                throw new Exception("Must save Outlook appointment before getting id");
            }
            return SetGoogleOutlookId(ga, oid);
        }

        internal static bool SetGoogleOutlookId(Event ga, string oid)
        {
            var key = OutlookPropertiesUtils.GetKey();

            if (ga.ExtendedProperties == null)
            {
                ga.ExtendedProperties = new Event.ExtendedPropertiesData();
            }

            if (ga.ExtendedProperties.Private__ == null)
            {
                ga.ExtendedProperties.Private__ = new Dictionary<string, string>();
            }

            // check if exists
            foreach (var p in ga.ExtendedProperties.Private__)
            {
                if (p.Key == key)
                {
                    if (ga.ExtendedProperties.Private__[p.Key] != oid)
                    {
                        ga.ExtendedProperties.Private__[p.Key] = oid;
                        return true;
                    }
                    return false;
                }
            }

            //not found
            var prop = new KeyValuePair<string, string>(key, oid);
            ga.ExtendedProperties.Private__.Add(prop);
            return true;
        }

        internal static string GetGoogleOutlookId(Event ga)
        {
            var key = OutlookPropertiesUtils.GetKey();
            string ret = null;

            // get extended prop
            if (ga.ExtendedProperties != null)
            {
                //First try to get private property
                ret = GetExtendedProperty(ga.ExtendedProperties.Private__, key);
                
                if (ret != null)
                    return ret;
                else //Then try to get Shared property
                    return GetExtendedProperty(ga.ExtendedProperties.Shared, key);                
            }
            return null;
        }

        private static string GetExtendedProperty(IDictionary<string, string> properties, string key)
        {
            if (properties != null)
            {
                foreach (var p in properties)
                {
                    if (p.Key == key)
                    {
                        return p.Value;
                    }
                }
            }

            return null;
        }

        internal static void ResetGoogleOutlookId(Event ga)
        {
            var key = OutlookPropertiesUtils.GetKey();

            if (ga.ExtendedProperties != null)
            {
                try
                { //First try to remove Shared property
                    RemoveExtendedPropertie(ga.ExtendedProperties.Shared, key);
                }
                catch (Exception ex)
                {
                    Log.Verbose(ex, "Error removing ExtendedProperties.Shared (key " + key + ")");
                }
                //Then remove private property
                RemoveExtendedPropertie(ga.ExtendedProperties.Private__, key);

            }
            //return false;
        }

        private static void RemoveExtendedPropertie(IDictionary<string,string> properties, string key)
        {
            if (properties != null)
            {
                // get extended prop
                foreach (var p in properties)
                {
                    if (p.Key == key)
                    {
                        // remove 
                       properties.Remove(p);
                       return;
                    }
                }
            }
            //return false;
        }

        /// <summary>
        /// Sets the syncId of the Outlook Appointment and the last sync date. 
        /// Please assure to always call this function when saving OutlookItem
        /// </summary>
        /// <param name="sync"></param>
        /// <param name="oa"></param>
        /// <param name="e"></param>
        internal static bool SetOutlookGoogleId(Outlook.AppointmentItem oa, Event e)
        {
            if (e.Id == null)
            {
                throw new NullReferenceException("GoogleAppointment must have a valid Id");
            }

            Outlook.UserProperties userProps = null;
            try
            {
                userProps = oa.UserProperties;
                return OutlookPropertiesUtils.SetOutlookGoogleId(userProps, e.Id, e.ETag);
            }
            finally
            {
                if (userProps != null)
                {
                    Marshal.ReleaseComObject(userProps);
                }
            }
        }

        internal static DateTime? GetOutlookLastSync(Outlook.AppointmentItem oa)
        {
            Outlook.UserProperties up = null;
            try
            {
                up = oa.UserProperties;
                return OutlookPropertiesUtils.GetOutlookLastSync(up);
            }
            finally
            {
                if (up != null)
                {
                    Marshal.ReleaseComObject(up);
                }
            }
        }

        internal static string GetOutlookLastEtag(Outlook.AppointmentItem oa)
        {
            Outlook.UserProperties up = null;
            try
            {
                up = oa.UserProperties;
                return OutlookPropertiesUtils.GetOutlookLastEtag(up);
            }
            finally
            {
                if (up != null)
                {
                    Marshal.ReleaseComObject(up);
                }
            }
        }

        internal static string GetOutlookGoogleId(Outlook.AppointmentItem oa)
        {
            Outlook.UserProperties up = null;

            try
            {
                up = oa.UserProperties;
                return OutlookPropertiesUtils.GetOutlookPropertyValue<string>(Synchronizer.OutlookPropertyNameId, up, Outlook.OlFormatText.olFormatTextText);
            }
            finally
            {
                if (up != null)
                {
                    Marshal.ReleaseComObject(up);
                }
            }
        }

        internal static void ResetOutlookGoogleId(Outlook.AppointmentItem oa)
        {
            Outlook.UserProperties up = null;
            try
            {
                up = oa.UserProperties;
                OutlookPropertiesUtils.ResetOutlookGoogleId(up);
            }
            finally
            {
                if (up != null)
                {
                    Marshal.ReleaseComObject(up);
                }
            }
        }

        internal static DateTime GetOutlookLastUpdated(AppointmentMatch match)
        {
            var lastUpdated = match.OutlookAppointment.LastModificationTime;

            if (match.OutlookAppointment.IsRecurring)
            {
                //adding Outlook exception is not changing modification date of the parent
                //first scan exceptions
                Outlook.RecurrencePattern rp = null;
                Outlook.Exceptions exceptions = null;

                try
                {
                    rp = match.OutlookAppointment.GetRecurrence();
                    exceptions = rp.Exceptions;
                    if (exceptions != null)
                    {
                        for (var i = exceptions.Count; i > 0; i--)
                        {
                            Outlook.Exception e = null;
                            Outlook.AppointmentItem ola = null;
                            try
                            {
                                e = exceptions[i];
                                if (e.Deleted &&
                                    //ToDo: Check, if we Don't sync deleted instances, if they are in the past, otherwise they will be synchronized again and again and again
                                    AppointmentSync.InSyncPeriod(e.OriginalDate)) //In Sync Range?
                                    //e.OriginalDate >= DateTime.Now)  //Or only in future?
                                    
                                {
                                    //for deleted exception it is not possible to get
                                    //their modification date 
                                    //==> Update it always, to not overlook a cancelled/deleted recurrence exception
                                    //ToDo: Check, but only if the recurrence exception is within sync range, to not update it again and again (especially for old exceptions in the past)                                    
                                    //ToDo: Or only the ones in the future?
                                    lastUpdated = DateTime.Now;
                                    Log.Debug($"Deleted Outlook recurrence exception found in the future (original date {e.OriginalDate.ToString()}, trying to delete it again (if not yet done): {match.OutlookAppointment.ToLogString()}.");
                                    break;
                                }
                                else
                                {   //Maybe it is possible to get the deletion date via the e.AppointmentItem.LastModificationTime???
                                    try
                                    {
                                        ola = e.AppointmentItem;
                                    }
                                    catch
                                    {
                                    }
                                    if (ola != null)
                                    {
                                        if (ola.LastModificationTime > lastUpdated)
                                        {
                                            lastUpdated = ola.LastModificationTime;
                                        }
                                    }
                                }
                            }
                            finally
                            {
                                if (ola != null)
                                {
                                    Marshal.ReleaseComObject(ola);
                                }
                                if (e != null)
                                {
                                    Marshal.ReleaseComObject(e);
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
            }

            if (lastUpdated == null)
                lastUpdated = DateTime.MinValue;

            if (lastUpdated.Kind == DateTimeKind.Utc)
                lastUpdated = TimeZoneInfo.ConvertTimeFromUtc(lastUpdated, TimeZoneInfo.Local);

            lastUpdated = lastUpdated.AddSeconds(-lastUpdated.Second);
            return lastUpdated;
        }

        internal static DateTime GetGoogleLastUpdated(AppointmentMatch match, Synchronizer sync)
        {
            DateTime lastUpdated = DateTime.MinValue;
            var now = DateTime.Now;
            if (match.GoogleAppointment.UpdatedDateTimeOffset.HasValue)
            {
                lastUpdated = match.GoogleAppointment.UpdatedDateTimeOffset.Value.DateTime;
                //consider GoogleAppointmentExceptions, because if they are updated, the master appointment doesn't have a new Saved TimeStamp
                foreach (var ga in match.GoogleAppointmentExceptions)
                {
                    if (ga.UpdatedDateTimeOffset != null)
                    {
                        var lastUpdatedGoogleException =ga.UpdatedDateTimeOffset.Value;
                        if (lastUpdatedGoogleException > lastUpdated)
                        {
                            lastUpdated = lastUpdatedGoogleException.DateTime;
                        }
                    }
                    else//ga.UpdatedDateTimeOffset == null, happens for cancelled events
                    {
                        var f_ga = sync.GetGoogleAppointment(ga.Id);
                        if (f_ga.UpdatedDateTimeOffset != null)
                        {
                            var lastUpdatedGoogleException = f_ga.UpdatedDateTimeOffset.Value;
                            if (lastUpdatedGoogleException > lastUpdated)
                            {
                                lastUpdated = lastUpdatedGoogleException.DateTime;
                            }
                        }
                        else if (match.OutlookAppointment.IsRecurring && match.OutlookAppointment.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster)
                        {
                            Outlook.AppointmentItem oa = null;
                            Outlook.RecurrencePattern rp = null;
                            Outlook.Exceptions outlookExceptions = null;                            


                            try
                            {
                                rp = match.OutlookAppointment.GetRecurrence();
                                outlookExceptions = rp.Exceptions;
                                DateTime date = AppointmentSync.outlookDateInvalid;//set any invalid date
                                if (ga.OriginalStartTime != null)
                                {
                                    oa = GetOccurrence(ga.OriginalStartTime, rp, ref date, match.OutlookAppointment);

                                    if (oa != null && oa.MeetingStatus != Outlook.OlMeetingStatus.olMeetingCanceled && oa.MeetingStatus != Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
                                    {
                                        lastUpdated = now;
                                        break; //no need to search further, already newest date set
                                    }

                                    if (outlookExceptions != null && date != AppointmentSync.outlookDateInvalid)
                                    {
                                        bool found = false;
                                        for (var j = outlookExceptions.Count; j > 0; j--)
                                        {
                                            Outlook.Exception e = null;
                                            try
                                            {
                                                e = outlookExceptions[j];
                                                if (e.Deleted && e.OriginalDate.Date == date.Date)
                                                {
                                                    found = true; //found deleted counterpart
                                                    break;
                                                }
                                            }
                                            finally
                                            {
                                                if (e != null)
                                                {
                                                    Marshal.ReleaseComObject(e);
                                                }
                                            }
                                        }
                                        if (!found)
                                        {
                                            Log.Debug($"Deleted Google recurrence exception found without deleted Outlook counterpart (original date {date.ToString()}, trying to delete it again (if not yet done): {ga.ToLogString()}.");
                                            lastUpdated = DateTime.Now;
                                            break;
                                        }
                                    }
                                }

                            }
                            catch (Exception ex) when ((uint)ex.HResult == 0x80004005)
                            {
                                //most likely there is deleted Outlook exception
                                lastUpdated = now;
                                break; //no need to search further, already newest date set
                            }
                            finally
                            {
                                if (rp != null)
                                {
                                    Marshal.ReleaseComObject(rp);
                                    rp = null;
                                }
                                if (oa != null)
                                {
                                    Marshal.ReleaseComObject(oa);
                                    oa = null;
                                }
                                if (outlookExceptions != null)
                                {
                                    Marshal.ReleaseComObject(outlookExceptions);
                                    outlookExceptions = null;
                                }
                            }

                                
                            
                            
                        }
                    }
                }
            }
            else
            {
                //Maybe complete new Google Appointment, therefore now LastUpdate???
                lastUpdated = now;
            }

            if (lastUpdated == null)
                lastUpdated = DateTime.MinValue;

            if (lastUpdated.Kind == DateTimeKind.Utc)
                lastUpdated = TimeZoneInfo.ConvertTimeFromUtc(lastUpdated, TimeZoneInfo.Local);

            //lastUpdated = lastUpdated.AddSeconds(-lastUpdated.Second); //Not needed for Outlook Appointments, because appointments can hold seconds (contacts cannot)
            return lastUpdated;
        }

        internal static Outlook.AppointmentItem GetOccurrence(EventDateTime originalStartTime, Outlook.RecurrencePattern rp, ref DateTime date, Outlook.AppointmentItem oa)
        {
            Outlook.AppointmentItem oe = null;
            if (!string.IsNullOrEmpty(originalStartTime.Date))
            {
                date = DateTime.Parse(originalStartTime.Date);                
            }
            else if (originalStartTime.DateTimeDateTimeOffset != null)
            {
                date = originalStartTime.DateTimeDateTimeOffset.Value.DateTime;                
            }

            if (!string.IsNullOrEmpty(originalStartTime.Date) || originalStartTime.DateTimeDateTimeOffset != null)
            {
                //ToDo: Check, if we Don't sync deleted instances, if they are in the past, otherwise they will be synchronized again and again and again
                if (AppointmentSync.InSyncPeriod(date)) //In sync range?
                                                        //if (date >= now.Date) //or only in the future?
                {

                    try
                    {
                        //First try with OriginalDateTime (always 12am for deleted exceptions)
                        oe = rp.GetOccurrence(date);
                    }
                    catch (COMException ex) when ((uint)ex.HResult == 0x80004005)
                    {
                        try
                        {//Second try with OriginalDate without Time
                            oe = rp.GetOccurrence(date.Date);
                        }
                        catch (COMException ex2) when ((uint)ex2.HResult == 0x80004005)
                        {
                            try
                            {//Third try with OriginalDate+slave.Start.TimeOfDay (consider original TimeOfDay)
                                oe = rp.GetOccurrence(date.Add(oa.Start.TimeOfDay));
                            }
                            catch (COMException ex3) when ((uint)ex3.HResult == 0x80004005)
                            {
                                Log.Debug(ex3, $"Representing outlook occurrence also not found: {oa.ToLogString()} - {originalStartTime}: {ex3}"); ;
                            }
                        }
                    }                    
                }
            }

            return oe;
        }
    }
}
