using Google.Apis.Calendar.v3.Data;
using Serilog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal static class AppointmentsMatcher
    {


        public delegate void NotificationHandler(string message);
        public static event NotificationHandler NotificationReceived;

        private static bool MatchAppointmentsByProperties(Outlook.AppointmentItem oa, Event ga)
        {
            if (oa.Subject != ga.Summary)
            {
                return false;
            }

            //check if both are all day events
            if (ga.Start != null)
            {
                if (oa.AllDayEvent)
                {
                    if (!string.IsNullOrEmpty(ga.Start.Date) && DateTime.Parse(ga.Start.Date) == oa.Start)
                    {
                        return true;
                    }
                }
                else
                {
                    if (ga.Start.DateTimeDateTimeOffset != null && oa.Start == ga.Start.DateTimeDateTimeOffset.Value.DateTime)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Matches outlook and Google appointment by a) id b) properties.
        /// </summary>
        /// <param name="sync">Syncronizer instance</param>
        /// <returns>Returns a list of match pairs (outlook appointment + Google appointment) for all appointment. Those that weren't matche will have it's peer set to null</returns>
        public static List<AppointmentMatch> MatchAppointments(Synchronizer sync)
        {
            Log.Information("Matching Outlook and Google appointments...");
            var result = new List<AppointmentMatch>();

            //for each outlook appointment try to get Google appointment id from user properties
            //if no match - try to match by properties
            //if no match - create a new match pair without Google appointment. 
            var OutlookAppointmentsWithoutSyncId = new Collection<Outlook.AppointmentItem>();
            #region Match first all outlookAppointments by sync id
            for (var i = 1; i <= sync.OutlookAppointments.Count; i++)
            {
                var oa = sync.OutlookAppointments[i] as Outlook.AppointmentItem;

                if (!sync.IsOutlookAppointmentToBeProcessed(oa))
                {
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;

                    //Logged already within above function
                    //Log.Warning($"Skipped Outlook appointment because could not be processed for several reasons (see Debug log messages in log file):" + (oa==null?"<null>":oa.ToLogString()));

                    if (oa != null)
                    {
                        Marshal.ReleaseComObject(oa);
                    }
                    
                    continue;
                }                

                string gid = null;
                try
                {
                    gid = AppointmentPropertiesUtils.GetOutlookGoogleId(oa);
                }
                catch (Exception ex)
                {
                    Log.Warning($"Skipped Outlook appointment because error accessing the properties (GoogleId reference):" + (oa == null ? "<null>" : oa.ToLogString()));
                    Log.Debug(ex, "Exception");

                    if (oa != null)
                    {
                        Marshal.ReleaseComObject(oa);
                    }

                    continue;
                }

                //try to match this appointment to one of Google appointments   
                var s = $"Matching appointment {i} of {sync.OutlookAppointments.Count} by id: {oa.ToLogString()}";
                NotificationReceived?.Invoke(s);
                Log.Debug(s);            

                if (gid != null)
                {
                    var ga = sync.GetGoogleAppointmentById(gid);

                    if (ga != null && !ga.Status.Equals("cancelled"))
                    {
                        //we found a match by google id, that is not deleted or cancelled yet
                        var match = new AppointmentMatch(oa, null);
                        match.matchedById = true;
                        match.AddGoogleAppointment(ga);
                        result.Add(match);
                        sync.GoogleAppointments.Remove(ga);
                    }
                    else
                    {
                        OutlookAppointmentsWithoutSyncId.Add(oa);
                    }
                }
                else
                {
                    OutlookAppointmentsWithoutSyncId.Add(oa);
                }
            }
            #endregion
            #region Match the remaining appointments by properties

            for (var i = 0; i < OutlookAppointmentsWithoutSyncId.Count; i++)
            {
                var oa = OutlookAppointmentsWithoutSyncId[i];

                var s = $"Matching appointment {i + 1} of {OutlookAppointmentsWithoutSyncId.Count} by unique properties: {oa.ToLogString()}";
                NotificationReceived?.Invoke(s);
                Log.Debug(s);

                //no match found by id => match by subject/title
                //create a default match pair with just outlook appointment.
                var match = new AppointmentMatch(oa, null);

                //foreach Google appointment try to match and create a match pair if found some match(es)
                for (var j = sync.GoogleAppointments.Count - 1; j >= 0; j--)
                {
                    var ga = sync.GoogleAppointments[j];
                    // only match if there is a appointment targetBody, else
                    // a matching Google appointment will be created at each sync                
                    if (!ga.Status.Equals("cancelled") && MatchAppointmentsByProperties(oa, ga))
                    {
                        match.AddGoogleAppointment(ga);
                        sync.GoogleAppointments.Remove(ga);
                    }
                }

                if (match.GoogleAppointment == null)
                {
                    string action;
                    if (sync.SyncOption != SyncOption.OutlookToGoogleOnly)
                    {
                        var gid = AppointmentPropertiesUtils.GetOutlookGoogleId(match.OutlookAppointment);
                        action = gid != null ? "Delete from Outlook" : "Add to Google";
                    }
                    else
                    {
                        action = "Add to Google";
                    }
                    Log.Debug($"No match found for Outlook appointment ({match.OutlookAppointment.ToLogString()}) => {action}");
                }

                result.Add(match);
            }
            #endregion

            var googleAppointmentExceptions = new List<Event>();

            //for each Google appointment that's left (they will be nonmatched) create a new match pair without outlook appointment. 
            for (var i = 0; i < sync.GoogleAppointments.Count; i++)
            {
                var ga = sync.GoogleAppointments[i];

                var s = $"Adding new Google appointment {i + 1} of {sync.GoogleAppointments.Count} by unique properties: {ga.ToLogString()}";
                NotificationReceived?.Invoke(s);
                Log.Debug(s);

                if (ga.RecurringEventId != null)
                {
                    sync.SkippedCountNotMatches++;
                    googleAppointmentExceptions.Add(ga);
                }
                else if (ga.Status.Equals("cancelled"))
                {
                    Log.Debug($"Skipping Google appointment found because it is cancelled: {ga.ToLogString()}");
                    //sync.SkippedCount++;
                    //sync.SkippedCountNotMatches++;
                }
                else if (string.IsNullOrEmpty(ga.Summary) && (ga.Start == null || (!ga.Start.DateTimeDateTimeOffset.HasValue && ga.Start.Date == null)))
                {
                    // no title or time
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    Log.Warning($"Skipped Google appointment because no unique property found (Subject or StartDate): {ga.ToLogString()}");
                }
                else
                {
                    var syncId = AppointmentPropertiesUtils.GetGoogleOutlookId(ga);
                    if (sync.SyncOption == SyncOption.GoogleToOutlookOnly && !string.IsNullOrEmpty(syncId))
                    {
                        Log.Warning($"Skipped Google Appointment because no unique property found (Subject or StartDate) and SyncOption {sync.SyncOption}: {ga.ToLogString()}");
                    }
                    else
                    {
                        var action = !string.IsNullOrEmpty(syncId) ? "Delete from Google" : "Add to Outlook";

                        Log.Debug($"No match found for Google appointment ({ga.ToLogString()}) => {action}");
                        result.Add(new AppointmentMatch(null, ga));
                    }
                }
            }

            //for each Google appointment exception, assign to proper match
            for (var i = 0; i < googleAppointmentExceptions.Count; i++)
            {
                var ga = googleAppointmentExceptions[i];
                var s = $"Adding Google appointment exception {i + 1} of {googleAppointmentExceptions.Count} : {ga.ToLogString()}";
                NotificationReceived?.Invoke(s);
                Log.Debug(s);

                var found = false;
                foreach (var match in result)
                {
                    if (match.GoogleAppointment != null && ga.RecurringEventId.Equals(match.GoogleAppointment.Id))
                    {
                        match.GoogleAppointmentExceptions.Add(ga);
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    var log = ga.ToLogString();
                    Log.Debug($"No match found for Google appointment exception: {log}");
                    if (string.IsNullOrWhiteSpace(log))
                    {
                        ga.ToDebugLog();
                    }
                    
                }
            }

            return result;
        }

        public static void SyncAppointments(Synchronizer sync)
        {
            for (var i = 0; i < sync.Appointments.Count; i++)
            {
                var match = sync.Appointments[i];
                var s = $"Syncing appointment {i + 1} of {sync.Appointments.Count}: {match.ToLogString()}";
                NotificationReceived?.Invoke(s);
                Log.Debug(s);
                SyncAppointment(match, sync);
            }
        }

        public static int RecipientsCount(Outlook.AppointmentItem oa)
        {
            Outlook.Recipients recipients = null;

            try
            {
                recipients = oa.Recipients;
                if (recipients != null)
                {
                    return recipients.Count;
                }
                else
                {
                    return 0;
                }
            }
            finally
            {
                if (recipients != null)
                {
                    Marshal.ReleaseComObject(recipients);
                }
            }
        }

        private static void SyncAppointmentNoGoogle(AppointmentMatch match, Synchronizer sync)
        {
            if (sync.SyncOption == SyncOption.OutlookToGoogleOnly)
            {                
                sync.RecreateGoogleAppointment(ref match);
                return;
            }

            if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
            {
                sync.SkippedCount++;
                Log.Debug($"Outlook appointment not added to Google, because of SyncOption {sync.SyncOption}: {match.OutlookAppointment.ToLogString()}");
                return;
            }

            var gid = AppointmentPropertiesUtils.GetOutlookGoogleId(match.OutlookAppointment);
            if (!string.IsNullOrEmpty(gid))
            {
                if (!sync.SyncDelete)
                {
                    return;
                }
                else if (!sync.PromptDelete && RecipientsCount(match.OutlookAppointment) <= 1)
                {
                    sync.DeleteOutlookResolution = DeleteResolution.DeleteOutlook;
                }
                else if (sync.DeleteOutlookResolution != DeleteResolution.DeleteOutlookAlways &&
                         sync.DeleteOutlookResolution != DeleteResolution.KeepOutlookAlways)
                {
                    using (var r = new ConflictResolver())
                    {
                        sync.DeleteOutlookResolution = r.ResolveDelete(match.OutlookAppointment);
                    }
                }

                switch (sync.DeleteOutlookResolution)
                {
                    case DeleteResolution.KeepOutlook:
                    case DeleteResolution.KeepOutlookAlways:
                        AppointmentPropertiesUtils.ResetOutlookGoogleId(match.OutlookAppointment);
                        break;
                    case DeleteResolution.DeleteOutlook:
                    case DeleteResolution.DeleteOutlookAlways:

                        if (RecipientsCount(match.OutlookAppointment) > 1)
                        {
                            //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment                                
                            Log.Information($"Outlook Appointment not deleted, because multiple participants found, invitation maybe NOT sent by Google: {match.OutlookAppointment.ToLogString()}");
                            AppointmentPropertiesUtils.ResetOutlookGoogleId(match.OutlookAppointment);
                            break;
                        }
                        else
                        {
                            //Avoid recreating a GoogleAppointment already existing
                            //==> Delete this OutlookAppointment instead if previous match existed but no match exists anymore
                            return;
                        }

                    default:
                        throw new ApplicationException("Cancelled");
                }
            }

            //create a Google appointment from Outlook appointment
            sync.RecreateGoogleAppointment(ref match);
        }

        private static void SyncAppointmentNoOutlook(AppointmentMatch match, Synchronizer sync)
        {
            if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
            {                
                sync.RecreateOutlookAppointment(ref match);
                return;
            }

            if (sync.SyncOption == SyncOption.OutlookToGoogleOnly)
            {
                sync.SkippedCount++;
                Log.Debug($"Google appointment not added to Outlook, because of SyncOption {sync.SyncOption}: {match.GoogleAppointment.Summary}");
                return;
            }

            var oid = AppointmentPropertiesUtils.GetGoogleOutlookId(match.GoogleAppointment);
            if (!string.IsNullOrEmpty(oid))
            {
                if (!sync.SyncDelete)
                {
                    return;
                }
                else if (!sync.PromptDelete)
                {
                    sync.DeleteGoogleResolution = DeleteResolution.DeleteGoogleAlways;
                }
                else if (sync.DeleteGoogleResolution != DeleteResolution.DeleteGoogleAlways &&
                         sync.DeleteGoogleResolution != DeleteResolution.KeepGoogleAlways)
                {
                    using (var r = new ConflictResolver())
                    {
                        sync.DeleteGoogleResolution = r.ResolveDelete(match.GoogleAppointment);
                    }
                }
                switch (sync.DeleteGoogleResolution)
                {
                    case DeleteResolution.KeepGoogle:
                    case DeleteResolution.KeepGoogleAlways:
                        AppointmentPropertiesUtils.ResetGoogleOutlookId(match.GoogleAppointment);
                        break;
                    case DeleteResolution.DeleteGoogle:
                    case DeleteResolution.DeleteGoogleAlways:
                        //Avoid recreating a OutlookAppointment already existing
                        //==> Delete this googleAppointment instead if previous match existed but no match exists anymore 
                        return;
                    default:
                        throw new ApplicationException("Cancelled");
                }
            }
            //create a Outlook appointment from Google appointment
            sync.RecreateOutlookAppointment(ref match);
        }




        private static void SyncAppointmentOutlookAndGoogle(AppointmentMatch match, Synchronizer sync)
        {
            //ToDo: Check how to overcome appointment recurrences, which need more than 60 seconds to update and therefore get updated again and again because of time tolerance 60 seconds violated again and again

            //merge appointment details                

            //determine if this appointment pair were synchronized
            //DateTime? lastUpdated = GetOutlookPropertyValueDateTime(match.OutlookAppointment, sync.OutlookPropertyNameUpdated);
            var lastSynced = AppointmentPropertiesUtils.GetOutlookLastSync(match.OutlookAppointment);
            if (lastSynced.HasValue)
            {
                //appointment pair was syncronysed before.

                //determine if Google appointment was updated since last sync                
                //lastSynced is stored without seconds. take that into account.
                var lastUpdatedOutlook = AppointmentPropertiesUtils.GetOutlookLastUpdated(match);
                var lastUpdatedGoogle = AppointmentPropertiesUtils.GetGoogleLastUpdated(match, sync);
                var OutlookUpdatedSinceLastSync = Utilities.UpdatedSinceLastSync(lastUpdatedOutlook, lastSynced.Value);
                var GoogleUpdatedSinceLastSync = Utilities.UpdatedSinceLastSync(lastUpdatedGoogle, lastSynced.Value);

                //ToDo: Too many updates, check if we can use eTag
                //if (!GoogleUpdatedSinceLastSync)
                //{
                    //var etagOutlook = AppointmentPropertiesUtils.GetOutlookLastEtag(match.OutlookAppointment);
                    //var eTagGoogle = match.GoogleAppointment.ETag;              
                    //if (!string.IsNullOrEmpty(etagOutlook) && !String.IsNullOrEmpty(eTagGoogle) && etagOutlook != eTagGoogle)
                    //   GoogleUpdatedSinceLastSync = true;
                //}

                //check if both outlok and Google appointments where updated sync last sync
                if (OutlookUpdatedSinceLastSync && GoogleUpdatedSinceLastSync)
                {
                    switch (sync.SyncOption)
                    {
                        case SyncOption.MergeOutlookWins:
                        case SyncOption.OutlookToGoogleOnly:
                            //overwrite Google appointment
                            Log.Debug($"Outlook and Google appointment have been updated, Outlook appointment is overwriting Google because of SyncOption {sync.SyncOption}: {match.OutlookAppointment.ToLogString()}.");
                            sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment, ref match.GoogleAppointmentExceptions);
                            break;
                        case SyncOption.MergeGoogleWins:
                        case SyncOption.GoogleToOutlookOnly:
                            //overwrite outlook appointment
                            Log.Debug($"Outlook and Google appointment have been updated, Google appointment is overwriting Outlook because of SyncOption {sync.SyncOption}: {match.GoogleAppointment.Summary}.");
                            sync.UpdateAppointment(ref match.GoogleAppointment, ref match.OutlookAppointment, match.GoogleAppointmentExceptions);
                            break;
                        case SyncOption.MergePrompt:
                            if (RecipientsCount(match.OutlookAppointment) > 1) //Only ask, if not multiple OutlookAppointments, because then it would anyway always update from Outlook to Google
                                sync.ConflictResolution = ConflictResolution.OutlookWins;
                            //promp for sync option
                            else if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.SkipAlways)                                
                            {
                                using (var r = new ConflictResolver())
                                {
                                    sync.ConflictResolution = r.Resolve(match.OutlookAppointment, match.GoogleAppointment, false);
                                }
                            }
                            switch (sync.ConflictResolution)
                            {
                                case ConflictResolution.Skip:
                                case ConflictResolution.SkipAlways:
                                    Log.Information($"User skipped appointment ({match}).");
                                    sync.SkippedCount++;
                                    break;
                                case ConflictResolution.OutlookWins:
                                case ConflictResolution.OutlookWinsAlways:
                                    sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment, ref match.GoogleAppointmentExceptions);
                                    break;
                                case ConflictResolution.GoogleWins:
                                case ConflictResolution.GoogleWinsAlways:
                                    sync.UpdateAppointment(ref match.GoogleAppointment, ref match.OutlookAppointment, match.GoogleAppointmentExceptions);
                                    break;
                                default:
                                    throw new ApplicationException("Cancelled");
                            }
                            break;
                    }
                    return;
                }

                //check if Outlook appointment was updated (with X second tolerance)
                if (sync.SyncOption != SyncOption.GoogleToOutlookOnly)
                {
                    //Outlook appointment was changed or changed Google appointment will be overwritten
                    if (sync.SyncOption == SyncOption.OutlookToGoogleOnly && GoogleUpdatedSinceLastSync)
                    {
                        Log.Debug($"Google appointment has been updated since last sync, but Outlook appointment is overwriting Google because of SyncOption {sync.SyncOption}: {match.OutlookAppointment.ToLogString()}.");
                        sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment, ref match.GoogleAppointmentExceptions);
                        return;
                    }
                    else if (OutlookUpdatedSinceLastSync)
                    {
                        sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment, ref match.GoogleAppointmentExceptions);
                        return;
                    }
                    //at the moment use Outlook as "master" source of appointments - in the event of a conflict Google appointment will be overwritten.
                    //TODO: control conflict resolution by SyncOption
                }

                //check if Google appointment was updated (with X second tolerance)
                if (sync.SyncOption != SyncOption.OutlookToGoogleOnly)
                {
                    //google appointment was changed or changed Outlook appointment will be overwritten
                    if (sync.SyncOption == SyncOption.GoogleToOutlookOnly && OutlookUpdatedSinceLastSync)
                    {
                        Log.Debug($"Outlook appointment has been updated since last sync, but Google appointment is overwriting Outlook because of SyncOption {sync.SyncOption}: {match.OutlookAppointment.ToLogString()}.");
                        sync.UpdateAppointment(ref match.GoogleAppointment, ref match.OutlookAppointment, match.GoogleAppointmentExceptions);
                        return;
                    }
                    else if (GoogleUpdatedSinceLastSync)
                    {
                        sync.UpdateAppointment(ref match.GoogleAppointment, ref match.OutlookAppointment, match.GoogleAppointmentExceptions);
                        return;
                    }
                }
            }
            else
            {
                //appointments were never synced.
                //merge appointments.
                switch (sync.SyncOption)
                {
                    case SyncOption.MergeOutlookWins:
                    case SyncOption.OutlookToGoogleOnly:
                        //overwrite Google appointment
                        sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment, ref match.GoogleAppointmentExceptions);
                        break;
                    case SyncOption.MergeGoogleWins:
                    case SyncOption.GoogleToOutlookOnly:
                        //overwrite outlook appointment
                        sync.UpdateAppointment(ref match.GoogleAppointment, ref match.OutlookAppointment, match.GoogleAppointmentExceptions);
                        break;
                    case SyncOption.MergePrompt:
                        //promp for sync option
                        if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                            sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.SkipAlways)
                        {
                            using (var r = new ConflictResolver())
                            {
                                sync.ConflictResolution = r.Resolve(match.OutlookAppointment, match.GoogleAppointment, true);
                            }
                        }
                        switch (sync.ConflictResolution)
                        {
                            case ConflictResolution.Skip:
                            case ConflictResolution.SkipAlways: //Keep both, Google AND Outlook
                                sync.Appointments.Add(new AppointmentMatch(match.OutlookAppointment, null));
                                sync.Appointments.Add(new AppointmentMatch(null, match.GoogleAppointment));
                                break;
                            case ConflictResolution.OutlookWins:
                            case ConflictResolution.OutlookWinsAlways:
                                sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment, ref match.GoogleAppointmentExceptions);
                                break;
                            case ConflictResolution.GoogleWins:
                            case ConflictResolution.GoogleWinsAlways:
                                sync.UpdateAppointment(ref match.GoogleAppointment, ref match.OutlookAppointment, match.GoogleAppointmentExceptions);
                                break;
                            default:
                                throw new ApplicationException("Canceled");
                        }
                        break;
                }
            }
        }

        public static void SyncAppointment(AppointmentMatch match, Synchronizer sync)
        {
            if (match.GoogleAppointment == null && match.OutlookAppointment != null)
            {
                //first check if maybe Google appointment was shifted to be outside sync range and
                //that is why we have not found it
                var gid = AppointmentPropertiesUtils.GetOutlookGoogleId(match.OutlookAppointment);
                if (!string.IsNullOrEmpty(gid))
                {
                    var ge = sync.GetGoogleAppointment(gid);
                    if (ge != null && ge.Status != "cancelled")
                    {
                        match.GoogleAppointment = ge;
                        SyncAppointmentOutlookAndGoogle(match, sync);
                        return;
                    }
                }

                //No Google appointment
                SyncAppointmentNoGoogle(match, sync);
            }
            else if (match.OutlookAppointment == null && match.GoogleAppointment != null)
            {
                //first check if maybe Outlook appointment was shifted to be outside sync range and
                //that is why we have not found it
                var oid = AppointmentPropertiesUtils.GetGoogleOutlookId(match.GoogleAppointment);
                if (!string.IsNullOrEmpty(oid))
                {
                    var oa = GetOutlookItem(sync, oid);
                    if (oa != null)
                    {
                        match.OutlookAppointment = oa;
                        SyncAppointmentOutlookAndGoogle(match, sync);
                        return;
                    }
                }

                //no Outlook appointment                               
                SyncAppointmentNoOutlook(match, sync);
            }
            else if (match.OutlookAppointment != null && match.GoogleAppointment != null)
            {
                SyncAppointmentOutlookAndGoogle(match, sync);
            }
        }

        private static Outlook.AppointmentItem GetOutlookItem(Synchronizer sync, string oid)
        {
            Outlook.AppointmentItem oa;
            try
            {
                dynamic item = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
                oa = item as Outlook.AppointmentItem;
                if (oa != null)
                {
                    /*if (oa.IsDeleted())
                    {
                        oa = null;
                    }*/
                    if (string.IsNullOrEmpty(oa.Subject) && oa.Start == AppointmentSync.outlookDateInvalid)
                    {
                        oa = null;
                    }
                    if (oa.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled || oa.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
                    {
                        oa = null;
                    }

                    Outlook.MAPIFolder fld = null;
                    Outlook.MAPIFolder s = null;
                    try
                    {
                        fld = oa.Parent as Outlook.MAPIFolder;
                        if (fld != null)
                        {
                            s = sync.GetAppoimentsFolder();
                            if (fld.FolderPath != s.FolderPath)
                            {
                                oa = null;
                            }
                        }
                        else
                        {
                            oa = null;
                        }
                    }
                    finally
                    {
                        if (fld != null)
                        {
                            Marshal.ReleaseComObject(fld);
                        }
                        if (s != null)
                        {
                            Marshal.ReleaseComObject(s);
                        }
                    }
                }
            }
            catch
            {
                oa = null;
            }

            return oa;
        }
    }

    internal class AppointmentMatch
    {
        public Outlook.AppointmentItem OutlookAppointment;
        public Event GoogleAppointment;
        public readonly List<Event> AllGoogleAppointmentMatches = new List<Event>(1);
        public Event LastGoogleAppointment;
        public List<Event> GoogleAppointmentExceptions = new List<Event>();
        public bool matchedById = false;

        //public bool GoogleAppointmentDirty;

        public AppointmentMatch(Outlook.AppointmentItem oa, Event ga)
        {
            OutlookAppointment = oa;
            GoogleAppointment = ga;
        }

        public string ToLogString()
        {
            var name = string.Empty;
            if (OutlookAppointment != null)
            {
                name = OutlookAppointment.ToLogString();
            }
            else if (GoogleAppointment != null)
            {
                name = GoogleAppointment.ToLogString();
            }
            return name;
        }

        public void AddGoogleAppointment(Event ga)
        {
            if (ga == null)
            {
                return;
            }
            //throw new ArgumentNullException("googleAppointment must not be null.");

            if (GoogleAppointment == null)
            {
                GoogleAppointment = ga;
            }

            //this to avoid searching the entire collection. 
            //if last appointment it what we are trying to add the we have already added it earlier
            if (LastGoogleAppointment == ga)
            {
                return;
            }

            if (!AllGoogleAppointmentMatches.Contains(ga))
            {
                AllGoogleAppointmentMatches.Add(ga);
            }

            LastGoogleAppointment = ga;
        }
    }
}
