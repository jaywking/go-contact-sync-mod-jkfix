using Google.Apis.PeopleService.v1.Data;
using Serilog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal static class ContactsMatcher
    { 

        public delegate void NotificationHandler(string message);
        public static event NotificationHandler NotificationReceived;

        private static void MatchContactsById(Synchronizer sync, List<ContactMatch> result, List<string> skippedOutlookIds, Collection<OutlookContactInfo> outlookContactsWithoutOutlookGoogleId)
        {
            for (var i = 1; i <= sync.OutlookContacts.Count; i++)
            {
                Outlook.ContactItem olc;
                try
                {
                    olc = sync.OutlookContacts[i] as Outlook.ContactItem;
                    if (olc == null)
                    {
                        if (sync.OutlookContacts[i] is Outlook.DistListItem)
                        {
                            Log.Debug("Skipping distribution list");
                            sync.TotalCount--;
                        }
                        else
                        {
                            dynamic item = sync.OutlookContacts[i];
                            var olClass = (Outlook.OlObjectClass)item.Class;
                            Log.Debug($"Empty Outlook contact found. Skipping ({olClass})");
                            sync.SkippedCount++;
                            sync.SkippedCountNotMatches++;
                        }
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some contacts throw exceptions
                    Log.Warning($"Accessing Outlook contact threw and exception. Skipping: {ex.Message}");
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    continue;
                }

                // sometimes contacts throw Exception when accessing their properties, so we give it a controlled try first.
                try
                {
                    var email1Address = olc.Email1Address;
                }
                catch (Exception ex)
                {
                    var message = string.Empty;
                    try
                    {
                        message = $"{message} {olc.ToLogString()}.";
                        //remember skippedOutlookIds to later not delete them if found on Google side
                        skippedOutlookIds.Add(string.Copy(olc.EntryID));
                    }
                    catch
                    {
                        message = $"Can't access contact details for outlook contact, got {ex.GetType()} - '{ex.Message}'. Skipping";
                    }

                    Log.Warning(message);
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    continue;
                }

                if (!IsContactValid(olc))
                {
                    Log.Warning($"Invalid outlook contact ({olc.ToLogString()}). Skipping");
                    skippedOutlookIds.Add(string.Copy(olc.EntryID));
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    continue;
                }

                var s = $"Matching contact {i} of {sync.OutlookContacts.Count} by id: {olc.ToLogString()}";

                NotificationReceived?.Invoke(s);
                Log.Debug(s);

                //try to match this contact to one of google contacts
                var userProperties = olc.UserProperties;
                var idProp = userProperties[Synchronizer.OutlookPropertyNameId];

                if (idProp == null)
                {
                    //accessing user properties by using [] is case sensitive, but later calling up.Add fails to add new property 
                    //as it is case insensitive.  As workaround first remove all properties that are equal if we ignore the case
                    OutlookPropertiesUtils.FindAndUnifySimilarProperty(userProperties, Synchronizer.OutlookPropertyNameSynced, Outlook.OlFormatDateTime.olFormatDateTimeBestFit);
                    idProp = OutlookPropertiesUtils.FindAndUnifySimilarProperty(userProperties, Synchronizer.OutlookPropertyNameId, Outlook.OlFormatText.olFormatTextText);

                    if (idProp != null)
                    {
                        olc.Save();
                    }
                }
                else
                {
                    var p = OutlookPropertiesUtils.FindAndUnifySimilarProperty(userProperties, Synchronizer.OutlookPropertyNameSynced, Outlook.OlFormatDateTime.olFormatDateTimeBestFit);

                    if (p != null)
                    {
                        olc.Save();
                    }
                }

                // Create our own info object to go into collections/lists, so we can free the Outlook objects and not run out of resources / exceed policy limits.
                var olci = new OutlookContactInfo(olc, sync);

                if (idProp != null)
                {
                    var gid = string.Copy((string)idProp.Value);
                    var foundContact = sync.GetGoogleContact(gid);
                    var match = new ContactMatch(olci, null);

                    //Check first, that this is not a duplicate 
                    //e.g. by copying an existing Outlook contact
                    //or by Outlook checked this as duplicate, but the user selected "Add new"
                    var duplicates = sync.OutlookContactByProperty(Synchronizer.OutlookPropertyNameId, gid);
                    if (duplicates.Count > 1)
                    {
                        Log.Debug("duplicates.Count > 1");

                        foreach (var duplicate in duplicates)
                        {
                            if (!string.IsNullOrEmpty(gid))
                            {
                                Log.Warning($"Duplicate Outlook contact found, resetting Outlook match and trying to match again: {duplicate.FileAs}");
                                var item = duplicate.GetOriginalItemFromOutlook();
                                try
                                {
                                    ContactPropertiesUtils.ResetOutlookGoogleId(sync, item);
                                    item.Save();
                                }
                                finally
                                {
                                    if (item != null)
                                    {
                                        Marshal.ReleaseComObject(item);
                                        item = null;
                                    }
                                }
                            }
                        }
                        if (foundContact != null && (!foundContact.Metadata.Deleted.HasValue || !foundContact.Metadata.Deleted.Value))
                        {
                            Log.Warning($"Duplicate Outlook contact found, resetting Google match and trying to match again: {foundContact.ToLogString()}");
                            ContactPropertiesUtils.ResetGoogleOutlookId(foundContact);
                        }
                        outlookContactsWithoutOutlookGoogleId.Add(olci);
                    }
                    else
                    {
                        if (foundContact != null && (!foundContact.Metadata.Deleted.HasValue || !foundContact.Metadata.Deleted.Value))
                        {
                            //we found a match by google id, that is not deleted yet
                            match.matchedById = true;
                            match.AddGoogleContact(foundContact);
                            result.Add(match);
                            //Remove the contact from the list to not sync it twice
                            sync.GoogleContacts.Remove(foundContact);
                        }
                        else
                        {
                            outlookContactsWithoutOutlookGoogleId.Add(olci);
                        }
                    }
                }
                else
                {
                    outlookContactsWithoutOutlookGoogleId.Add(olci);
                }
            }
        }

        private static bool IsMatching(OutlookContactInfo olci, Person entry)
        {
            // only match if there is either an email or telephone or else
            // a matching google contact will be created at each sync
            //1. try to match by FileAs
            //1.1 try to match by FullName
            //2. try to match by primary email
            //3. try to match by mobile phone number, don't match by home or business numbers, because several people may share the same home or business number
            //4. try to math Company, if Google Title is null, i.e. the contact doesn't have a name and title, only a company
            var entryTitleFirstLastAndSuffix = ContactPropertiesUtils.GetGoogleTitleFirstLastAndSuffix(entry);
            var fileAsValue = ContactPropertiesUtils.GetGoogleFileAsValue(entry);
            var unstructuredName = ContactPropertiesUtils.GetGoogleUnstructuredName(entry);            
            var companyName = ContactPropertiesUtils.GetGooglePrimaryOrganizationName(entry);           
            var emailValue = ContactPropertiesUtils.GetGooglePrimaryEmailValue(entry);           

            if ((!string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(fileAsValue) && olci.FileAs.Trim().Equals(fileAsValue.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase)) ||  //Replace twice to not replace a \r\n by \r\r\n. This is necessary because \r\n are saved as \n only to google
                (!string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(unstructuredName) && olci.FileAs.Trim().Equals(unstructuredName.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase)) ||
                (!string.IsNullOrEmpty(olci.FullName) && !string.IsNullOrEmpty(unstructuredName) && olci.FullName.Trim().Equals(unstructuredName.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase)) ||
                (!string.IsNullOrEmpty(olci.TitleFirstLastAndSuffix) && !string.IsNullOrEmpty(entryTitleFirstLastAndSuffix) && olci.TitleFirstLastAndSuffix.Trim().Equals(entryTitleFirstLastAndSuffix.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase)) ||
                (!string.IsNullOrEmpty(olci.Email1Address) && !string.IsNullOrEmpty(emailValue) && olci.Email1Address.Trim().Equals(emailValue, StringComparison.InvariantCultureIgnoreCase)) ||
                (olci.MobileTelephoneNumber != null && FindPhone(olci.MobileTelephoneNumber, entry.PhoneNumbers) != null) ||
                (!string.IsNullOrEmpty(olci.FileAs) && string.IsNullOrEmpty(fileAsValue) && !string.IsNullOrEmpty(companyName) && olci.FileAs.Trim().Equals(companyName, StringComparison.InvariantCultureIgnoreCase))
                )
            {
                return true;
            }
            return false;
        }

        private static void MatchContactsByProperties(Synchronizer sync, List<ContactMatch> result, Collection<OutlookContactInfo> outlookContactsWithoutOutlookGoogleId, out DuplicateDataException duplicatesFound)
        {
            var duplicateGoogleMatches = string.Empty;
            var duplicateOutlookContacts = string.Empty;

            for (var i = 0; i < outlookContactsWithoutOutlookGoogleId.Count; i++)
            {
                var olci = outlookContactsWithoutOutlookGoogleId[i];
                var s = $"Matching contact {i + 1} of {outlookContactsWithoutOutlookGoogleId.Count} by unique properties: {olci}";

                NotificationReceived?.Invoke(s);
                Log.Debug(s);

                //no match found by id => match by common properties
                //create a default match pair with just outlook contact.
                var match = new ContactMatch(olci, null);

                //for each google contact try to match and create a match pair if found some match(es)
                for (var j = sync.GoogleContacts.Count - 1; j >= 0; j--)
                {
                    var entry = sync.GoogleContacts[j];
                    if (entry.Metadata.Deleted.HasValue && entry.Metadata.Deleted.Value)
                    {
                        continue;
                    }
                    if (IsMatching(olci, entry))
                    {
                        Log.Debug("Match with google contact found");
                        match.AddGoogleContact(entry);
                        sync.GoogleContacts.Remove(entry);
                    }
                }

                if (match.AllGoogleContactMatches == null || match.AllGoogleContactMatches.Count == 0)
                {
                    //Check, if this Outlook contact has a match in the google duplicates
                    var duplicateFound = false;
                    foreach (var duplicate in sync.GoogleContactDuplicates)
                    {
                        if (IsMatching(olci, duplicate.AllGoogleContactMatches[0]))
                        {
                            duplicateFound = true;
                            Log.Debug("Duplicate found");
                            duplicate.AddOutlookContact(olci);
                            sync.OutlookContactDuplicates.Add(match);
                            if (string.IsNullOrEmpty(duplicateOutlookContacts))
                            {
                                duplicateOutlookContacts = "Outlook contact found that has been already identified as duplicate Google contact (either same email, Mobile or FullName) and cannot be synchronized. Please delete or resolve duplicates of:";
                            }

                            var str = $"{olci.FileAs} ({olci.Email1Address}, {olci.MobileTelephoneNumber})";
                            if (!duplicateOutlookContacts.Contains(str))
                            {
                                duplicateOutlookContacts += Environment.NewLine + str;
                            }

                            break;
                        }
                    }

                    if (!duplicateFound)
                    {                       
                        var gid = olci.UserProperties.GoogleContactId;
                        var action = ((sync.SyncOption == SyncOption.OutlookToGoogleOnly) || string.IsNullOrEmpty(gid)) ? "Add to Google" : "Delete from Outlook";
                        Log.Debug($"No match found for Outlook contact ({olci.FileAs}) => {action}");
                    }
                }
                else
                {
                    //Remember Google duplicates to later react to it when resetting matches or syncing
                    //ResetMatches: Also reset the duplicates
                    //Sync: Skip duplicates (don't sync duplicates to be fail safe)
                    if (match.AllGoogleContactMatches.Count > 1)
                    {
                        sync.GoogleContactDuplicates.Add(match);
                        foreach (var entry in match.AllGoogleContactMatches)
                        {
                            //Create message for duplicatesFound exception
                            if (string.IsNullOrEmpty(duplicateGoogleMatches))
                            {
                                duplicateGoogleMatches = "Outlook contacts matching with multiple Google contacts have been found (either same email, Mobile, FullName or company) and cannot be synchronized. Please delete or resolve duplicates of:";
                            }

                            var str = $"{olci.FileAs} ({olci.Email1Address}, {olci.MobileTelephoneNumber})";
                            if (!duplicateGoogleMatches.Contains(str))
                            {
                                duplicateGoogleMatches += Environment.NewLine + str;
                            }
                        }
                    }
                }

                result.Add(match);
            }

            duplicatesFound = !string.IsNullOrEmpty(duplicateGoogleMatches) || !string.IsNullOrEmpty(duplicateOutlookContacts)
                 ? new DuplicateDataException(duplicateGoogleMatches + Environment.NewLine + Environment.NewLine + duplicateOutlookContacts)
                 : null;
        }

        /// <summary>
        /// Matches outlook and google contact by a) google id b) properties.
        /// </summary>
        /// <param name="sync">Syncronizer instance</param>
        /// <param name="duplicatesFound">Exception returned, if duplicates have been found (null else)</param>
        /// <returns>Returns a list of match pairs (outlook contact + google contact) for all contact. Those that weren't matche will have it's peer set to null</returns>
        public static List<ContactMatch> MatchContacts(Synchronizer sync, out DuplicateDataException duplicatesFound)
        {
            Log.Information("Matching Outlook and Google contacts...");
            var result = new List<ContactMatch>();

            sync.GoogleContactDuplicates = new Collection<ContactMatch>();
            sync.OutlookContactDuplicates = new Collection<ContactMatch>();

            var skippedOutlookIds = new List<string>();

            //for each outlook contact try to get google contact id from user properties
            //if no match - try to match by properties
            //if no match - create a new match pair without google contact. 

            var outlookContactsWithoutOutlookGoogleId = new Collection<OutlookContactInfo>();

            MatchContactsById(sync, result, skippedOutlookIds, outlookContactsWithoutOutlookGoogleId);

            MatchContactsByProperties(sync, result, outlookContactsWithoutOutlookGoogleId, out duplicatesFound);

            //for each google contact that's left (they will be nonmatched) create a new match pair without outlook contact. 
            for (var i = 0; i < sync.GoogleContacts.Count; i++)
            {

                var entry = sync.GoogleContacts[i];
                var googleUniqueIdentifierName = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(entry);
                NotificationReceived?.Invoke($"Adding new Google contact {i + 1} of {sync.GoogleContacts.Count} by unique properties: {googleUniqueIdentifierName} ...");

                // only match if there is either an email or mobile phone or a name or a company
                // otherwise a matching google contact will be created at each sync
                /*var mobileExists = false;
                if (entry.PhoneNumbers != null)
                {
                    foreach (var phone in entry.PhoneNumbers)
                    {
                        if (phone != null && phone.Type != null && phone.Type.Equals(ContactSync.PHONE_MOBILE, StringComparison.InvariantCultureIgnoreCase)) //ToDo: Get proper enum
                        {
                            mobileExists = true;
                            break;
                        }
                    }
                }*/

                var googleOutlookId = ContactPropertiesUtils.GetGoogleOutlookContactId(entry);                
                if (!string.IsNullOrEmpty(googleOutlookId) && skippedOutlookIds.Contains(googleOutlookId))
                {
                    Log.Warning($"Skipped GoogleContact because Outlook contact couldn't be matched because of previous problem (see log): {googleUniqueIdentifierName}");
                }
                //else if (entry.EmailAddresses.Count == 0 && !mobileExists && string.IsNullOrEmpty(fileAs.Value) && string.IsNullOrEmpty(name.UnstructuredName) && (entry.Organizations.Count == 0 || string.IsNullOrEmpty(entry.Organizations[0].Name)))
                else if (!string.IsNullOrEmpty(googleOutlookId)
                     && sync.OutlookContacts != null && sync.OutlookContacts.Count > 0 
                    && string.IsNullOrEmpty(ContactPropertiesUtils.GetGooglePrimaryEmailValue(entry)) 
                    && string.IsNullOrEmpty(ContactPropertiesUtils.GetGooglePrimaryPhoneValue(entry)) 
                    && string.IsNullOrEmpty(ContactPropertiesUtils.GetGoogleFileAsValue(entry)) 
                    & string.IsNullOrEmpty(ContactPropertiesUtils.GetGoogleUnstructuredName(entry)) 
                    && string.IsNullOrEmpty(ContactPropertiesUtils.GetGoogleTitleFirstLastAndSuffix(entry)) 
                    && string.IsNullOrEmpty(ContactPropertiesUtils.GetGooglePrimaryOrganizationName(entry)))
                {
                    // no unique identifier found, e.g. like mobile or email or UnstructuredName or Company

                    if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        Log.Warning($"Skipped GoogleContact because no unique property found (Email1 or mobile or name or company) and SyncOption {sync.SyncOption}: {ContactMatch.GetSummary(entry)}");
                    }
                    else
                    {
                        //ToDo: For now I use the ResolveDelete function, because it is almost the same, maybe we introduce a separate function for this ans also include DeleteGoogleAlways checkbox
                        using (var r = new ConflictResolver())
                        {
                            var res = r.ResolveDelete(entry);

                            if (res == DeleteResolution.DeleteGoogle || res == DeleteResolution.DeleteGoogleAlways)
                            {
                                ContactPropertiesUtils.SetGoogleOutlookContactId(entry, "-1"); //just set a dummy Id to delete this entry later on
                                Log.Debug($"No match found for Google contact ({entry.ToLogString()}) => {res.ToString()}");
                                //sync.SaveContact(new ContactMatch(null, entry));
                                result.Add(new ContactMatch(null, entry));
                            }
                            else
                            {
                                sync.SkippedCount++;
                                sync.SkippedCountNotMatches++;

                                Log.Warning($"Skipped GoogleContact because no unique property found (Email1 or mobile or name or company): {ContactMatch.GetSummary(entry)}");
                            }
                        }
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(googleOutlookId) 
                     && sync.OutlookContacts != null && sync.OutlookContacts.Count > 0 
                     && sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        Log.Warning($"Skipped GoogleContact because no unique property found (Email1 or mobile or name or company) and SyncOption {sync.SyncOption}: {ContactMatch.GetSummary(entry)}");
                    }
                    else
                    {
                        var action = !string.IsNullOrEmpty(googleOutlookId) ? "Delete from Google" : "Add to Outlook";
                        Log.Debug($"No match found for Google contact ({entry.ToLogString()}) => {action}");
                        var match = new ContactMatch(null, entry);
                        result.Add(match);
                    }
                }
            }
            return result;
        }

        private static bool IsContactValid(Outlook.ContactItem contact)
        {
            if (!string.IsNullOrEmpty(contact.Email1Address))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.Email2Address))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.Email3Address))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.HomeTelephoneNumber))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.BusinessTelephoneNumber))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.MobileTelephoneNumber))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.HomeAddress))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.BusinessAddress))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.OtherAddress))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(contact.Body))
            {
                return true;
            }

            return contact.Birthday != DateTime.MinValue;
        }

        public static void SyncContacts(Synchronizer sync)
        {
            for (var i = 0; i < sync.Contacts.Count; i++)
            {
                var match = sync.Contacts[i];
                var s = $"Syncing contact {i + 1} of {sync.Contacts.Count}: {match}";
                NotificationReceived?.Invoke(s);
                Log.Debug(s);
                SyncContact(match, sync);
            }
        }

        private static void SyncContactNoGoogle(Outlook.ContactItem outlookContactItem, ContactMatch match, Synchronizer sync)
        {
            //no google contact                               
            var gid = match.OutlookContact.UserProperties.GoogleContactId;
            var askDelete = false;

            if (sync.SyncOption == SyncOption.GoogleToOutlookOnly && string.IsNullOrEmpty(gid))
            {
                askDelete = true;
                //sync.SkippedCount++;
                Log.Debug($"Outlook Contact not added to Google, because of SyncOption {sync.SyncOption}: {match.OutlookContact.FileAs}");
                //return;
            }
            else if (!string.IsNullOrEmpty(gid))
            {
                //Redundant check if exist, but in case an error occurred in MatchContacts
                askDelete = true;
                var matchingGoogleContact = sync.GetGoogleContact(gid);
                if (matchingGoogleContact == null)
                {
                    if (!sync.SyncDelete)
                    {
                        return;//ToDo: Check: kept on OutlookSide? and skip logged for not deleting?
                    }
                    else if (!sync.PromptDelete)
                    {
                        //sync.DeleteOutlookResolution = DeleteResolution.DeleteOutlookAlways;
                        return;//==> Delete this outlookContact instead if previous match existed but no match exists anymore
                    }
                    else if (sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        //sync.DeleteGoogleResolution = DeleteResolution.KeepOutlook;
                        askDelete = false;
                    }
                }
            }


            if (askDelete)
            {
                if (sync.DeleteOutlookResolution != DeleteResolution.DeleteOutlookAlways &&
                    sync.DeleteOutlookResolution != DeleteResolution.KeepOutlookAlways)
                {
                    using (var r = new ConflictResolver())
                    {
                        sync.DeleteOutlookResolution = r.ResolveDelete(match.OutlookContact);
                    }
                }

                switch (sync.DeleteOutlookResolution)
                {
                    case DeleteResolution.KeepOutlook:
                    case DeleteResolution.KeepOutlookAlways:
                        if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                            return; //Don't Delete this outlookContact, but also not create a Google contact
                        else
                        {
                            ContactPropertiesUtils.ResetOutlookGoogleId(sync, outlookContactItem);
                            break; //Recreate Google Contact from Outlook
                        }
                    case DeleteResolution.DeleteOutlook:
                    case DeleteResolution.DeleteOutlookAlways:
                        //Avoid recreating a GoogleContact already existing
                        //==> Delete this outlookContact instead if previous match existed but no match exists anymore
                        return;
                    default:
                        throw new ApplicationException("Cancelled");
                }
            }
            

            //create a Google contact from Outlook contact
            match.GoogleContact = new Person();
            sync.UpdateContact(outlookContactItem, match.GoogleContact, match);
        }

        private static void SyncContactNoOutlook(ContactMatch match, Synchronizer sync)
        {
            // no outlook contact
            var outlookId = ContactPropertiesUtils.GetGoogleOutlookContactId(match.GoogleContact);
            var askDelete = false;

            if (sync.SyncOption == SyncOption.OutlookToGoogleOnly && string.IsNullOrEmpty(outlookId))
            {
                askDelete = true;
                //sync.SkippedCount++;
                Log.Debug($"Google Person not added to Outlook, because of SyncOption {sync.SyncOption}: {ContactPropertiesUtils.GetGoogleUniqueIdentifierName(match.GoogleContact)}");
                //return;
            }
            else if (!string.IsNullOrEmpty(outlookId))
            {
                askDelete = true;
                if (!sync.SyncDelete)
                {
                    return; //ToDo: Check: kept on GoogleSide? and skip logged for not deleting?
                }
                else if (!sync.PromptDelete)
                {
                    //sync.DeleteGoogleResolution = DeleteResolution.DeleteGoogleAlways;
                    return; //==> Delete this googleContact instead if previous match existed but no match exists anymore    
                }
                else if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                {
                    //sync.DeleteGoogleResolution = DeleteResolution.KeepGoogle;
                    askDelete = false;
                }
            }

            if (askDelete)
            {
                if (sync.DeleteGoogleResolution != DeleteResolution.DeleteGoogleAlways &&
                    sync.DeleteGoogleResolution != DeleteResolution.KeepGoogleAlways)
                {
                    using (var r = new ConflictResolver())
                    {
                        sync.DeleteGoogleResolution = r.ResolveDelete(match.GoogleContact);
                    }
                }

                switch (sync.DeleteGoogleResolution)
                {
                    case DeleteResolution.KeepGoogle:
                    case DeleteResolution.KeepGoogleAlways:
                        if (sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                            return; //Don't Delete this googleContact, but also not create an Outlook contact
                        else
                        {
                            ContactPropertiesUtils.ResetGoogleOutlookId(match.GoogleContact);
                            break; //Recreate Outlook Contact from Google
                        }
                    case DeleteResolution.DeleteGoogle:
                    case DeleteResolution.DeleteGoogleAlways:
                        //Avoid recreating a OutlookContact already existing
                        //==> Delete this googleContact instead if previous match existed but no match exists anymore                
                        return;
                    default:
                        throw new ApplicationException("Cancelled");
                }
            }

            //create a Outlook contact from Google contact                                                            
            var outlookContactItem = Synchronizer.CreateOutlookContactItem(Synchronizer.SyncContactsFolder);

            sync.UpdateContact(match.GoogleContact, outlookContactItem, match.GoogleContactDirty, match.matchedById);
            match.OutlookContact = new OutlookContactInfo(outlookContactItem, sync);
        }

        
        private static void SyncContactOutlookAndGoogle(Outlook.ContactItem oc, ContactMatch match, Synchronizer sync)
        {
            //merge contact details                

            //determine if this contact pair were synchronized
            //lastSynced is stored without seconds. take that into account.
            var lastSynced = match.OutlookContact.UserProperties.LastSync;
            if (lastSynced.HasValue)
            {
                //contact pair was syncronysed before.
                //determine if Outlook and Google contact were updated since last sync                
                var lastUpdatedOutlook = ContactPropertiesUtils.GetOutlookLastUpdated(match.OutlookContact);
                var lastUpdatedGoogle = ContactPropertiesUtils.GetGoogleLastUpdated(match.GoogleContact);

                var OutlookUpdatedSinceLastSync = Utilities.UpdatedSinceLastSync(lastUpdatedOutlook, lastSynced.Value);
                var GoogleUpdatedSinceLastSync = Utilities.UpdatedSinceLastSync(lastUpdatedGoogle, lastSynced.Value);

                //ToDo: Too many updates, check if we can use eTag
                //if (!GoogleUpdatedSinceLastSync)
                //{
                //    var etagOutlook = ContactPropertiesUtils.GetOutlookLastEtag(match.OutlookContact.GetOriginalItemFromOutlook());
                //    var eTagGoogle = match.GoogleContact.ETag;
                //    if (!string.IsNullOrEmpty(etagOutlook) && !String.IsNullOrEmpty(eTagGoogle) && etagOutlook != eTagGoogle)
                //        GoogleUpdatedSinceLastSync = true;
                //}

                //check if both outlok and google contacts where updated sync last sync
                if (OutlookUpdatedSinceLastSync && GoogleUpdatedSinceLastSync)
                {
                    //both contacts were updated.
                    //options: 1) ignore 2) lose one based on SyncOption
                    //throw new Exception("Both contacts were updated!");
                    switch (sync.SyncOption)
                    {
                        case SyncOption.MergeOutlookWins:
                        case SyncOption.OutlookToGoogleOnly:
                            //overwrite google contact
                            Log.Debug($"Outlook and Google contact have been updated, Outlook contact is overwriting Google because of SyncOption {sync.SyncOption}: {match.OutlookContact.FileAs}.");
                            sync.UpdateContact(oc, match.GoogleContact, match);
                            break;
                        case SyncOption.MergeGoogleWins:
                        case SyncOption.GoogleToOutlookOnly:
                            //overwrite outlook contact
                            Log.Debug($"Outlook and Google contact have been updated, Google contact is overwriting Outlook because of SyncOption {sync.SyncOption}: {match.OutlookContact.FileAs}.");
                            sync.UpdateContact(match.GoogleContact, oc, match.GoogleContactDirty, match.matchedById);
                            break;
                        case SyncOption.MergePrompt:
                            //promp for sync option
                            Log.Debug($"Merge: {match.OutlookContact.FileAs}. Outlook updated: {match.OutlookContact.LastModificationTime}, Google updated: {ContactPropertiesUtils.GetGoogleLastUpdated(match.GoogleContact)}");
                            if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.SkipAlways)
                            {
                                using (var r = new ConflictResolver())
                                {
                                    sync.ConflictResolution = r.Resolve(match, false);
                                }
                            }
                            switch (sync.ConflictResolution)
                            {
                                case ConflictResolution.Skip:
                                case ConflictResolution.SkipAlways:
                                    Log.Information($"User skipped contact ({match}).");
                                    sync.SkippedCount++;
                                    break;
                                case ConflictResolution.OutlookWins:
                                case ConflictResolution.OutlookWinsAlways:
                                    sync.UpdateContact(oc, match.GoogleContact, match);
                                    break;
                                case ConflictResolution.GoogleWins:
                                case ConflictResolution.GoogleWinsAlways:
                                    sync.UpdateContact(match.GoogleContact, oc, match.GoogleContactDirty, match.matchedById);
                                    break;
                                default:
                                    throw new ApplicationException("Cancelled");
                            }
                            break;
                    }
                    return;
                }

                //check if outlook contact was updated (with X second tolerance)
                if (sync.SyncOption != SyncOption.GoogleToOutlookOnly)
                {
                    //outlook contact was changed or changed Google contact will be overwritten
                    if (sync.SyncOption == SyncOption.OutlookToGoogleOnly && GoogleUpdatedSinceLastSync)
                    {
                        Log.Debug($"Google contact has been updated since last sync, but Outlook contact is overwriting Google because of SyncOption {sync.SyncOption}: {match.OutlookContact.FileAs}.");
                        sync.UpdateContact(oc, match.GoogleContact, match);
                        return;
                    }
                    else if (OutlookUpdatedSinceLastSync)
                    {
                        sync.UpdateContact(oc, match.GoogleContact, match);
                        return;
                    }
                    //at the moment use outlook as "master" source of contacts - in the event of a conflict google contact will be overwritten.
                    //TODO: control conflict resolution by SyncOption
                }

                //check if google contact was updated (with X second tolerance)
                if (sync.SyncOption != SyncOption.OutlookToGoogleOnly)
                {
                    //google contact was changed or changed Outlook contact will be overwritten
                    if (sync.SyncOption == SyncOption.GoogleToOutlookOnly && OutlookUpdatedSinceLastSync)
                    {
                        Log.Debug($"Outlook contact has been updated since last sync, but Google contact is overwriting Outlook because of SyncOption {sync.SyncOption}: {match.OutlookContact.FileAs}.");
                        sync.UpdateContact(match.GoogleContact, oc, match.GoogleContactDirty, match.matchedById);
                        return;
                    }
                    else if (GoogleUpdatedSinceLastSync)
                    {
                        sync.UpdateContact(match.GoogleContact, oc, match.GoogleContactDirty, match.matchedById);
                        return;
                    }
                }
            }
            else
            {
                //contacts were never synced.
                //merge contacts.
                switch (sync.SyncOption)
                {
                    case SyncOption.MergeOutlookWins:
                    case SyncOption.OutlookToGoogleOnly:
                        //overwrite google contact
                        sync.UpdateContact(oc, match.GoogleContact, match);
                        break;
                    case SyncOption.MergeGoogleWins:
                    case SyncOption.GoogleToOutlookOnly:
                        //overwrite outlook contact
                        sync.UpdateContact(match.GoogleContact, oc, match.GoogleContactDirty, match.matchedById);
                        break;
                    case SyncOption.MergePrompt:
                        //promp for sync option
                        if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                            sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                            sync.ConflictResolution != ConflictResolution.SkipAlways)
                        {
                            using (var r = new ConflictResolver())
                            {
                                sync.ConflictResolution = r.Resolve(match, true);
                            }
                        }
                        switch (sync.ConflictResolution)
                        {
                            case ConflictResolution.Skip:
                            case ConflictResolution.SkipAlways: //Keep both, Google AND Outlook
                                sync.Contacts.Add(new ContactMatch(match.OutlookContact, null));
                                sync.Contacts.Add(new ContactMatch(null, match.GoogleContact));
                                break;
                            case ConflictResolution.OutlookWins:
                            case ConflictResolution.OutlookWinsAlways:
                                sync.UpdateContact(oc, match.GoogleContact, match);
                                break;
                            case ConflictResolution.GoogleWins:
                            case ConflictResolution.GoogleWinsAlways:
                                sync.UpdateContact(match.GoogleContact, oc, match.GoogleContactDirty, match.matchedById);
                                break;
                            default:
                                throw new ApplicationException("Cancelled");
                        }
                        break;
                }
            }
        }



        public static void SyncContact(ContactMatch match, Synchronizer sync)
        {
            var oc = match.OutlookContact?.GetOriginalItemFromOutlook();

            try
            {
                if (match.GoogleContact == null && match.OutlookContact != null)
                {
                    SyncContactNoGoogle(oc, match, sync);
                }
                else if (match.OutlookContact == null && match.GoogleContact != null)
                {
                    SyncContactNoOutlook(match, sync);
                }
                else if (match.OutlookContact != null && match.GoogleContact != null)
                {
                    SyncContactOutlookAndGoogle(oc, match, sync);
                }
                else
                {
                    throw new ArgumentNullException("ContactMatch has all peers null.");
                }
            }
            catch (ArgumentNullException)
            {
                throw;
            }
            catch (Exception e)
            {
                throw new Exception($"Error syncing contact {(match.OutlookContact != null ? match.OutlookContact.FileAs : ContactPropertiesUtils.GetGoogleUniqueIdentifierName(match.GoogleContact))}: {e.Message}", e);
            }
            finally
            {
                if (oc != null && match.OutlookContact != null)
                {
                    match.OutlookContact.Update(oc, sync);
                    Marshal.ReleaseComObject(oc);
                }
            }
        }

        private static PhoneNumber FindPhone(string number, IList<PhoneNumber> phones)
        {
            if (string.IsNullOrEmpty(number))
            {
                return null;
            }

            if (phones == null)
            {
                return null;
            }

            foreach (var phone in phones)
            {
                if (phone != null && phone.Value != null && number.Trim().Equals(phone.Value.Trim(), StringComparison.InvariantCultureIgnoreCase))
                {
                    return phone;
                }
            }

            return null;
        }

        /// <summary>
        /// Adds new Google Groups to the Google account.
        /// </summary>
        /// <param name="sync"></param>
        public static void SyncGroups(Synchronizer sync)
        {
            foreach (var match in sync.Contacts)
            {
                if (match.OutlookContact != null && !string.IsNullOrEmpty(match.OutlookContact.Categories))
                {
                    var cats = Utilities.GetOutlookGroups(match.OutlookContact.Categories);
                    foreach (var cat in cats)
                    {
                        if (cat != null)
                        {   //Contact group name "Starred in Android" is a reserved legacy name, was used in old Contact API. 
                            //Backward compliancy by using the system "contactGroups/starred" group instead
                            if (cat.Trim().Equals("Starred in Android", StringComparison.InvariantCultureIgnoreCase))
                            {
                                match.OutlookContact.Categories = match.OutlookContact.Categories.Replace("Starred in Android", "starred");
                            }
                            else
                            {
                                var g = sync.GetGoogleGroupByName(cat);
                                if (g == null)
                                {
                                    // create group                            
                                    g = sync.CreateGroup(cat);
                                    g = sync.SaveGoogleGroup(g);
                                    sync.GoogleGroups.Add(g);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    internal class ContactMatch
    {
        public OutlookContactInfo OutlookContact;
        public Person GoogleContact;
        public readonly List<Person> AllGoogleContactMatches = new List<Person>();
        public readonly List<OutlookContactInfo> AllOutlookContactMatches = new List<OutlookContactInfo>();
        public bool matchedById = false;

        public bool GoogleContactDirty;

        public ContactMatch(OutlookContactInfo outlookContact, Person googleContact)
        {
            AddOutlookContact(outlookContact);
            AddGoogleContact(googleContact);
        }

        public void AddGoogleContact(Person googleContact)
        {
            if (googleContact == null)
            {
                return;
            }

            if (GoogleContact == null)
            {
                GoogleContact = googleContact;
            }

            //this to avoid searching the entire collection. 
            //if last contact it what we are trying to add the we have already added it earlier
            //if (LastGoogleContact == googleContact)
            //    return;

            if (!AllGoogleContactMatches.Contains(googleContact))
            {
                AllGoogleContactMatches.Add(googleContact);
            }

            //LastGoogleContact = googleContact;
        }

        public void AddOutlookContact(OutlookContactInfo outlookContact)
        {
            if (outlookContact == null)
            {
                return;
            }
            //throw new ArgumentNullException("outlookContact must not be null.");

            if (OutlookContact == null)
            {
                OutlookContact = outlookContact;
            }

            //this to avoid searching the entire collection. 
            //if last contact it what we are trying to add the we have already added it earlier
            //if (LastGoogleContact == googleContact)
            //    return;

            if (!AllOutlookContactMatches.Contains(outlookContact))
            {
                AllOutlookContactMatches.Add(outlookContact);
            }

            //LastGoogleContact = googleContact;
        }

        public override string ToString()
        {
            if (OutlookContact != null)
            {
                var s = OutlookContact.ToString();
                if (!string.IsNullOrWhiteSpace(s))
                {
                    return s;
                }
            }

            if (GoogleContact != null)
            {
                return GetName(GoogleContact);
            }
            return string.Empty;
        }

        public static string GetName(Person gc)
        {            
            var name = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
            }

            var googleContactName = ContactPropertiesUtils.GetGooglePrimaryName(gc);
            if (googleContactName != null)
            {
                name = googleContactName.UnstructuredName;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                }
            }

            if (gc.Organizations != null && gc.Organizations.Count > 0)
            {
                name = gc.Organizations[0].Name;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                }
            }

            if (gc.EmailAddresses != null && gc.EmailAddresses.Count > 0)
            {
                name = gc.EmailAddresses[0].Value;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                }
            }
            return string.Empty;
        }

        public static string GetSummary(Outlook.ContactItem outlookContact)
        {
            var name = OutlookContactInfo.GetTitleFirstLastAndSuffix(outlookContact);
            var summary = string.Empty;

            if (!string.IsNullOrEmpty(name))
            {
                summary += "Name: " + name.Trim().Replace("  ", " ") + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.FirstName))
            {
                summary += "First Name: " + outlookContact.FirstName + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.LastName))
            {
                summary += "Last Name: " + outlookContact.LastName + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.CompanyName))
            {
                summary += "Company: " + outlookContact.CompanyName + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.Department))
            {
                summary += "Department: " + outlookContact.Department + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.Email1Address))
            {
                summary += "Email1: " + outlookContact.Email1Address + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.Email2Address))
            {
                summary += "Email2: " + outlookContact.Email2Address + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.Email3Address))
            {
                summary += "Email3: " + outlookContact.Email3Address + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.MobileTelephoneNumber))
            {
                summary += "Mobile Phone: " + outlookContact.MobileTelephoneNumber + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.HomeTelephoneNumber))
            {
                summary += "Home Phone: " + outlookContact.HomeTelephoneNumber + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.Home2TelephoneNumber))
            {
                summary += "Home Phone2: " + outlookContact.Home2TelephoneNumber + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.BusinessTelephoneNumber))
            {
                summary += "Business Phone: " + outlookContact.BusinessTelephoneNumber + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.Business2TelephoneNumber))
            {
                summary += "Business Phone2: " + outlookContact.Business2TelephoneNumber + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.OtherTelephoneNumber))
            {
                summary += "Other Phone: " + outlookContact.OtherTelephoneNumber + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.HomeAddress))
            {
                summary += "Home Address: " + outlookContact.HomeAddress + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.BusinessAddress))
            {
                summary += "Business Address: " + outlookContact.BusinessAddress + "\r\n";
            }

            if (!string.IsNullOrEmpty(outlookContact.OtherAddress))
            {
                summary += "Other Address: " + outlookContact.OtherAddress + "\r\n";
            }

            return summary;
        }

        public static string GetSummary(Person googleContact)
        {
            var name = ContactPropertiesUtils.GetGoogleTitleFirstLastAndSuffix(googleContact);
            var summary = string.Empty;

            if (!string.IsNullOrEmpty(name))
            {
                summary += "Name: " + name.Trim().Replace("  ", " ") + "\r\n";
            }

            var googleContactName = ContactPropertiesUtils.GetGooglePrimaryName(googleContact);
            if (googleContactName != null && !string.IsNullOrEmpty(googleContactName.GivenName))
            {
                summary += "First Name: " + googleContactName.GivenName + "\r\n";
            }

            if (googleContactName != null && !string.IsNullOrEmpty(googleContactName.FamilyName))
            {
                summary += "Last Name: " + googleContactName.FamilyName + "\r\n";
            }

            if (googleContact.Organizations != null)
                for (var i = 0; i < googleContact.Organizations.Count; i++)
                {
                    var company = googleContact.Organizations[i].Name;
                    var department = googleContact.Organizations[i].Department;
                    if (!string.IsNullOrEmpty(company))
                    {
                        summary += "Company: " + company + "\r\n";
                    }
                    if (!string.IsNullOrEmpty(department))
                    {
                        summary += "Department: " + department + "\r\n";
                    }
                }

            if (googleContact.EmailAddresses != null)
                for (var i = 0; i < googleContact.EmailAddresses.Count; i++)
                {
                    var email = googleContact.EmailAddresses[i].Value;
                    if (!string.IsNullOrEmpty(email))
                    {
                        summary += "Email" + (i + 1) + ": " + email + "\r\n";
                    }
                }

            if (googleContact.PhoneNumbers != null)
                foreach (var phone in googleContact.PhoneNumbers)
                {
                    if (phone != null && phone.Type != null && !string.IsNullOrEmpty(phone.Value))
                    {
                        if (phone.Type.Trim().Equals(ContactSync.PHONE_MOBILE, StringComparison.InvariantCultureIgnoreCase)) //ToDo: Get proper enum
                        {
                            summary += "Mobile Phone: ";
                        }

                        if (phone.Type.Trim().Equals(ContactSync.HOME, StringComparison.InvariantCultureIgnoreCase))
                        {
                            summary += "Home Phone: ";
                        }

                        if (phone.Type.Trim().Equals(ContactSync.WORK, StringComparison.InvariantCultureIgnoreCase))
                        {
                            summary += "Business Phone: ";
                        }

                        if (phone.Type.Trim().Equals(ContactSync.OTHER, StringComparison.InvariantCultureIgnoreCase) || phone.Type.Trim().Equals(ContactSync.ANDERE, StringComparison.InvariantCultureIgnoreCase))
                        {
                            summary += "Other Phone: ";
                        }

                        summary += phone.Value + "\r\n";
                    }
                }

            if (googleContact.Addresses != null)
                foreach (var address in googleContact.Addresses)
                {
                    if (address != null && address.Type != null && !string.IsNullOrEmpty(address.FormattedValue))
                    {
                        if (address.Type.Trim().Equals(ContactSync.HOME, StringComparison.InvariantCultureIgnoreCase)) //ToDo: Get proper enum
                        {
                            summary += "Home Address: ";
                        }

                        if (address.Type.Trim().Equals(ContactSync.WORK, StringComparison.InvariantCultureIgnoreCase))
                        {
                            summary += "Business Address: ";
                        }

                        if (address.Type.Trim().Equals(ContactSync.OTHER, StringComparison.InvariantCultureIgnoreCase) || address.Type.Trim().Equals(ContactSync.ANDERE, StringComparison.InvariantCultureIgnoreCase))
                        {
                            summary += "Other Address: ";
                        }

                        summary += address.FormattedValue + "\r\n";
                    }
                }
            return summary;
        }
    }
}