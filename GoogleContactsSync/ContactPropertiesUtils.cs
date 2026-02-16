using Google.Apis.PeopleService.v1.Data;
//using Google.Apis.People.v1.Data;

using Serilog;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal static class ContactPropertiesUtils
    {
        internal const string CONTACT = "CONTACT";
        internal const string INITIALS = "INITIALS";
        internal const string DEFAULT = "DEFAULT";

        internal static string GetOutlookId(Outlook.ContactItem oc)
        {
            return oc.EntryID;
        }

        internal static string GetGoogleId(Person gc)
        {
            //if (gc.ResourceName == null)
            //    throw new Exception("ResourceName is null");
            //else
            //    return gc.ResourceName.Replace("people/", string.Empty);
            //return gc.ResourceName;

            var googleContactEntry = GetGoogleContact(gc);
            if (googleContactEntry != null)
                return googleContactEntry.Id;

            return null;

        }

        internal static DateTime GetGoogleLastUpdated(Person gc)
        {
            var lastUpdated = DateTime.MinValue;
            var googleContactEntry = GetGoogleContact(gc);
            if (googleContactEntry != null && googleContactEntry.UpdateTimeDateTimeOffset.HasValue)
            {
                if (googleContactEntry.UpdateTimeDateTimeOffset is DateTimeOffset time)
                    lastUpdated = time.DateTime;
                else
                    DateTime.TryParse(googleContactEntry.UpdateTimeDateTimeOffset.ToString(),out lastUpdated);                
            }

            if (lastUpdated.Kind == DateTimeKind.Utc)
                lastUpdated = TimeZoneInfo.ConvertTimeFromUtc(lastUpdated, TimeZoneInfo.Local);
             
            return lastUpdated;
        }

        internal static DateTime GetOutlookLastUpdated(OutlookContactInfo oci)
        {
            var lastUpdated = DateTime.MinValue;
            if (oci.LastModificationTime != null)
                lastUpdated = oci.LastModificationTime;

            if (lastUpdated.Kind == DateTimeKind.Utc)
                lastUpdated = TimeZoneInfo.ConvertTimeFromUtc(lastUpdated, TimeZoneInfo.Local);

            return lastUpdated;
        }

        internal static string GetGoogleInitialsValue(Person gc)
        {
            var initials = GetGoogleInitials(gc);
            return initials == null ? string.Empty : initials.Value;
        }

        internal static Nickname GetGoogleInitials(Person gc)
        {
            return GetGoogleNickname(gc, INITIALS); //ToDo: Check, find proper enum
        }

        internal static string GetGoogleNickNameValue(Person gc)
        {
            var nick = GetGoogleNickName(gc);
            return nick == null ? string.Empty : nick.Value;
        }

        internal static Nickname GetGoogleNickName(Person gc)
        {
            return GetGoogleNickname(gc, null); //ToDo: Check, find proper enum, why not DEFAULT???
        }

        private static Nickname GetGoogleNickname(Person gc, string type)
        {
            if (gc != null && gc.Nicknames != null)
            {
                foreach (var nickname in gc.Nicknames)
                {
                    if (nickname != null && nickname.Metadata != null && nickname.Metadata.Source != null && nickname.Metadata.Source.Type != null 
                        && nickname.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase) && nickname.Type == type) //ToDo: Check
                        return nickname;
                }
            }

            return null;
        }

        internal static EmailAddress GetGooglePrimaryEmail(Person gc)
        {

            if (gc != null && gc.EmailAddresses != null && gc.EmailAddresses.Count > 0)
            {
                foreach (var email in gc.EmailAddresses)
                {
                    if (email != null && email.Metadata != null && email.Metadata.Source != null && email.Metadata.Source.Type != null
                        && email.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase) 
                        && (email.Metadata.Primary??false)) //ToDo: Check
                        return email;
                }
                foreach (var email in gc.EmailAddresses)
                {
                    if (email != null && !string.IsNullOrWhiteSpace(email.Value)) //ToDo: Check
                        return email;
                }
                return gc.EmailAddresses[0];
            }

            return null;

        }

        internal static string GetGooglePrimaryEmailValue(Person gc)
        {
            var email = GetGooglePrimaryEmail(gc);
            return email == null ? string.Empty : email.Value;
        }

        internal static Name GetGooglePrimaryName(Person gc)
        {

            if (gc != null && gc.Names != null && gc.Names.Count > 0)
            {
                foreach (var name in gc.Names)
                {
                    if (name != null && name.Metadata != null && name.Metadata.Source != null && name.Metadata.Source.Type != null
                        && name.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase)
                        && (name.Metadata.Primary ?? false)) //ToDo: Check
                        return name;
                }
                foreach (var name in gc.Names)
                {
                    if (name != null && (!string.IsNullOrWhiteSpace(name.UnstructuredName) || !string.IsNullOrWhiteSpace(GetGoogleTitleFirstLastAndSuffix(gc)))) //ToDo: Check
                        return name;
                }
                return gc.Names[0];
            }

            return null;

        }

        internal static string GetGoogleUnstructuredName(Person gc)
        {
            var name = GetGooglePrimaryName(gc);
            return name == null ? string.Empty : name.UnstructuredName;
        }

        internal static string GetGoogleTitleFirstLastAndSuffix(Person gc)
        {
            var name = ContactPropertiesUtils.GetGooglePrimaryName(gc);
            if (name == null)
                return string.Empty;
            else
                return GetTitleFirstLastAndSuffix(name.HonorificPrefix, name.GivenName, name.MiddleName, name.FamilyName, name.HonorificSuffix);
        }

        internal static PhoneNumber GetGooglePrimaryPhone(Person gc, string type)
        {

            if (gc != null && gc.PhoneNumbers != null && gc.PhoneNumbers.Count > 0)
            {
                foreach (var phone in gc.PhoneNumbers)
                {
                    if (phone != null && phone.Metadata != null && phone.Metadata.Source != null && phone.Metadata.Source.Type != null
                        && phone.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase)
                        && (phone.Metadata.Primary ?? false)
                        && (string.IsNullOrEmpty(type) || type.Equals(phone.Type))) //ToDo: Check
                        return phone;
                }
                foreach (var phone in gc.PhoneNumbers)
                {
                    if (phone != null && !string.IsNullOrWhiteSpace(phone.Value) && (string.IsNullOrEmpty(type) || type.Equals(phone.Type)))
                        return phone;
                }
                var phoneDefault = gc.PhoneNumbers[0];
                if (string.IsNullOrEmpty(type) || type.Equals(phoneDefault.Type))
                    return phoneDefault;
            }

            return null;

        }

        internal static string GetGooglePrimaryPhoneValue(Person gc)
        {
            var phone = GetGooglePrimaryPhone(gc, null);
            if (phone == null || string.IsNullOrEmpty(phone.Value))
                phone = GetGooglePrimaryPhone(gc, ContactSync.PHONE_MOBILE);
            return phone == null ? string.Empty : phone.Value;
        }

        internal static string GetGoogleUniqueIdentifierName(Person gc)
        {
            Log.Verbose($" Starting GetGoogleUniqueIdentifierName(Person gc)...");
            Log.Verbose($"  Startinging GetGoogleFileAsValue...");
            var identifier = GetGoogleFileAsValue(gc);
            Log.Verbose($"  Finisheded GetGoogleFileAsValue: " + identifier);
            if (string.IsNullOrEmpty(identifier))
            {
                Log.Verbose($"  Starting GetGoogleTitleFirstLastAndSuffix(gc)...");
                identifier = GetGoogleTitleFirstLastAndSuffix(gc);
                Log.Verbose($"  Finished GetGoogleTitleFirstLastAndSuffix(gc): " + identifier);
            }
            if (string.IsNullOrEmpty(identifier))
            {
                Log.Verbose($"  Starting GetGooglePrimaryEmailValue(gc)...");
                identifier = GetGooglePrimaryEmailValue(gc);
                Log.Verbose($"  Finished GetGooglePrimaryEmailValue(gc): " + identifier);
            }
                
            if (string.IsNullOrEmpty(identifier))
            {
                Log.Verbose($"  Starting GetGooglePrimaryPhoneValue(gc)...");
                identifier = GetGooglePrimaryPhoneValue(gc);
                Log.Verbose($"  Finished GetGooglePrimaryPhoneValue(gc): " + identifier);
            }
                
            if (string.IsNullOrEmpty(identifier))
            {
                Log.Verbose($"  Starting GetGooglePrimaryOrganizationName(gc)...");
                identifier = GetGooglePrimaryOrganizationName(gc);
                Log.Verbose($"  Finished GetGooglePrimaryOrganizationName(gc): " + identifier);
            }

            /* Removing because of Endless-Loop, see
             * https://sourceforge.net/p/googlesyncmod/support-requests/838/
             * https://sourceforge.net/p/googlesyncmod/bugs/1285/
            if (string.IsNullOrEmpty(identifier))
            {
                Log.Verbose($"  Starting gc.ToLogString()...");
                identifier = gc.ToLogString();
                Log.Verbose($"  Finished gc.ToLogString(): " + identifier);
            }
            */

            if (string.IsNullOrEmpty(identifier))
            {
                Log.Verbose($"  Starting gc.ToString()...");
                identifier = gc.ToString();
                Log.Verbose($"  Finished gc.ToString(): " + identifier);
            }
                
            if (string.IsNullOrEmpty(identifier) || identifier.Contains(identifier.GetType().FullName))
            {
                Log.Verbose($"  Starting UnIdentified...");
                identifier = "UnIdentified";
                Log.Verbose($"  Finished UnIdentified: " + identifier);
            }
                
            Log.Debug($" Finished GetGoogleUniqueIdentifierName(Person gc): " + identifier);
            return identifier;

        }

        internal static string GetTitleFirstLastAndSuffix(string title, string firstName, string middleName, string lastName, string suffix)
        {
            string s;
            if (string.IsNullOrWhiteSpace(title))
            {
                s = firstName?.RemoveNewLines();
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(firstName))
                {
                    s = title.RemoveNewLines() + " " + firstName.RemoveNewLines();
                }
                else
                {
                    s = title.RemoveNewLines();
                }
            }

            if (string.IsNullOrWhiteSpace(s))
            {
                s = middleName.RemoveNewLines();
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(middleName))
                {
                    s = s + " " + middleName.RemoveNewLines();
                }
            }

            if (string.IsNullOrWhiteSpace(s))
            {
                s = lastName.RemoveNewLines();
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(lastName))
                {
                    s = s + " " + lastName.RemoveNewLines();
                }
            }

            if (string.IsNullOrWhiteSpace(s))
            {
                s = suffix.RemoveNewLines();
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(suffix))
                {
                    s = s + " " + suffix.RemoveNewLines();
                }
            }

            return s;
        }

        internal static FileAs GetGoogleFileAs(Person gc)
        {

            if (gc != null && gc.FileAses != null && gc.FileAses.Count > 0)
            {
                foreach (var fileAs in gc.FileAses)
                {
                    if (fileAs != null && fileAs.Metadata != null && fileAs.Metadata.Source != null && fileAs.Metadata.Source.Type != null
                        && fileAs.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase)
                        && (fileAs.Metadata.Primary ?? false)) //ToDo: Check
                        return fileAs;
                }

                return gc.FileAses[0];
            }

            return null;

        }

        internal static string GetGoogleFileAsValue(Person gc)
        {
            var fileAs = GetGoogleFileAs(gc);
            return fileAs == null?string.Empty:fileAs.Value;
        }

        internal static Biography GetGoogleBiography(Person gc)
        {

            if (gc != null && gc.Biographies != null && gc.Biographies.Count > 0)
            {
                foreach (var bio in gc.Biographies)
                {
                    if (bio != null && bio.Metadata != null && bio.Metadata.Source != null && bio.Metadata.Source.Type != null
                        && bio.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase)
                        && (bio.Metadata.Primary ?? false)) //ToDo: Check
                        return bio;
                }
                return gc.Biographies[0];
            }

            return null;

        }

        internal static string GetGoogleBiographyValue(Person gc)
        {
            var bio = GetGoogleBiography(gc);
            return bio == null ? string.Empty : bio.Value;
        }

        internal static Birthday GetGoogleBirthday(Person gc)
        {

            if (gc != null && gc.Birthdays != null && gc.Birthdays.Count > 0)
            {
                foreach (var birthday in gc.Birthdays )
                {
                    if (birthday != null && birthday.Metadata != null && birthday.Metadata.Source != null && birthday.Metadata.Source.Type != null
                        && birthday.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase)
                        && (birthday.Metadata.Primary ?? false)) //ToDo: Check
                        return birthday;
                }
                return gc.Birthdays[0];
            }

            return null;

        }
       

        internal static Organization GetGooglePrimaryOrganization(Person gc)
        {

            if (gc != null && gc.Organizations != null && gc.Organizations.Count > 0)
            {
                foreach (var org in gc.Organizations)
                {
                    if (org != null && org.Metadata != null && org.Metadata.Source != null && org.Metadata.Source.Type != null
                        && org.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase)
                        && (org.Metadata.Primary??false)) //ToDo: Check
                        return org;
                }
                foreach (var org in gc.Organizations)
                {
                    if (org != null && !string.IsNullOrWhiteSpace(org.Name))
                        return org;
                }
                return gc.Organizations[0];
            }

            return null;

        }

        internal static string GetGooglePrimaryOrganizationName(Person gc)
        {
            var org = GetGooglePrimaryOrganization(gc);
            return org==null?string.Empty:org.Name;
        }



        internal static UserDefined GetGoogleUserDefined(Person gc, string key)
        {
            UserDefined foundUserDefinedByKey = null;
            if (gc != null && gc.UserDefined != null)
            {
                foreach (var userDefined in gc.UserDefined)
                {
                    if (userDefined != null && userDefined.Metadata != null && userDefined.Metadata.Source != null && userDefined.Metadata.Source.Type != null
                        && userDefined.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase) && userDefined.Key == key) //ToDo: Check
                    {
                        if (!string.IsNullOrEmpty(userDefined.Value))
                            return userDefined;
                        else
                            foundUserDefinedByKey = userDefined;
                    }
                }
            }

            return foundUserDefinedByKey;

        }

        internal static Location GetGoogleOfficeLocation(Person gc)
        {
            if (gc != null && gc.Locations != null)
            {
                foreach (var location in gc.Locations)
                {
                    if (location != null && location.Metadata != null && location.Metadata.Source != null && location.Metadata.Source.Type != null
                        && location.Metadata.Source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase) && location.Type == ContactSync.DESK) //ToDo: Check, and get proper enum
                        return location;
                }
            }

            return null;

        }

        internal static string GetGoogleOfficeLocationValue(Person gc)
        {
            var loc = GetGoogleOfficeLocation(gc);
            return loc == null ? string.Empty : loc.Value;
        }

        private static Source GetGoogleContact(Person gc)
        {
            if (gc == null || gc.Metadata == null || gc.Metadata.Sources == null)
                //throw new Exception("Sources list is null, Google Contact Entry could not be extracted");
                return null;

            foreach (var source in gc.Metadata.Sources)
            {
                if (source != null && source.Type != null && source.Type.Equals(CONTACT, StringComparison.InvariantCultureIgnoreCase)) //ToDo: Find correct Enum
                    return source;
            }

            //throw new Exception("Sources list didn'T contain any Google Contact Entry");
            return null;

        }

        internal static void SetGoogleOutlookId(Person gc, Outlook.ContactItem oc)
        {
            var id = GetOutlookId(oc);
            if (id == null)
            {
                throw new Exception("Must save outlook contact before getting id");
            }
            SetGoogleOutlookContactId(gc, id);
        }

        internal static void SetGoogleOutlookContactId(Person gc, string oid)
        {
            // check if exists
            var found = false;
            var key = OutlookPropertiesUtils.GetKey();
            if (gc.ClientData == null)
                gc.ClientData = new List<ClientData>();
            foreach (var p in gc.ClientData)
            {
                if (p.Key == key)
                {
                    if (p.Value != oid)
                    {
                        p.Value = oid;
                    }
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                var prop = new ClientData()
                {
                    Key = key,
                    Value = oid
                };
                gc.ClientData.Add(prop);
            }
        }


        internal static string GetGoogleOutlookContactId(Person gc)
        {
            var key = OutlookPropertiesUtils.GetKey();

            // get extended prop
            if (gc.ClientData != null)
                foreach (var p in gc.ClientData)
                {
                    if (p.Key == key)
                        if (!string.IsNullOrEmpty(p.Value) && !p.Value.Trim().Equals(key,StringComparison.InvariantCultureIgnoreCase)) //ToDo: Bug in PeopleApi, if ClientData.Remove is not saved into Google and if set to null or String.Empty, it saves the Key into the Value
                            return p.Value;
                        else
                            return null;
                    
                }
            return null;
        }

        internal static void ResetGoogleOutlookId(Person gc)
        {
            var key = OutlookPropertiesUtils.GetKey();
            // get extended prop
            foreach (var p in gc.ClientData)
            {
                if (p.Key == key)
                {
                    // remove 
                    //gc.ClientData.Remove(p); //ToDo: Just remove doesn't work with new Google People Api
                    p.Value = string.Empty; //ToDo: Bug in PeopleApi, if ClientData.Remove is not saved into Google and if set to null or String.Empty, it saves the Key into the Value
                    return;
                }
            }
        }

        /// <summary>
        /// Sets the syncId of the Outlook contact and the last sync date. 
        /// Please assure to always call this function when saving OutlookItem
        /// </summary>
        /// <param name="sync"></param>
        /// <param name="oc"></param>
        /// <param name="gc"></param>
        internal static bool SetOutlookGoogleId(Outlook.ContactItem oc, Person gc)
        {
            var gid = ContactPropertiesUtils.GetGoogleId(gc);
            if (gid == null)
            {
                throw new NullReferenceException("GoogleContact must have a valid Id");
            }

            Outlook.UserProperties up = null;
            try
            {
                up = oc.UserProperties;
                return OutlookPropertiesUtils.SetOutlookGoogleId(up, gid, gc.ETag);
            }
            finally
            {
                if (up != null)
                {
                    Marshal.ReleaseComObject(up);
                }
            }
        }

        internal static string GetOutlookGoogleId(Outlook.ContactItem oc)
        {
            Outlook.UserProperties up = null;

            try
            {
                up = oc.UserProperties;
                var id = OutlookPropertiesUtils.GetOutlookPropertyValue<string>(Synchronizer.OutlookPropertyNameId, up, Outlook.OlFormatText.olFormatTextText);
                if (!string.IsNullOrEmpty(id))
                {
                    var slash = id.LastIndexOf("/"); //ToDo: For Backward compatibility with old GoogleContact-API: remove the prefix from the id (e.g. 1c0d39680d700698 from http://www.google.com/m8/feeds/contacts/saller.flo%40gmail.com/base/1c0d39680d700698)
                    id = id.Substring(slash + 1);
                }

                return id;
            }
            finally
            {
                if (up != null)
                {
                    Marshal.ReleaseComObject(up);
                }
            }
        }

        internal static DateTime? GetOutlookLastSync(Outlook.ContactItem oc)
        {
            Outlook.UserProperties up = null;

            try
            {
                up = oc.UserProperties;
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

        internal static string GetOutlookLastEtag(Outlook.ContactItem oc)
        {
            Outlook.UserProperties up = null;

            try
            {
                up = oc.UserProperties;
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

        internal static void ResetOutlookGoogleId(Synchronizer sync, Outlook.ContactItem olc)
        {
            Outlook.UserProperties up = null;

            try
            {
                up = olc.UserProperties;
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

        internal static string GetOutlookEmailAddress1(Outlook.ContactItem oc)
        {
            return GetOutlookEmailAddress(oc, oc.Email1AddressType, oc.Email1Address);
        }

        internal static string GetOutlookEmailAddress2(Outlook.ContactItem oc)
        {
            return GetOutlookEmailAddress(oc, oc.Email2AddressType, oc.Email2Address);
        }

        internal static string GetOutlookEmailAddress3(Outlook.ContactItem oc)
        {
            return GetOutlookEmailAddress(oc, oc.Email3AddressType, oc.Email3Address);
        }

        private static string GetOutlookEmailAddress(Outlook.ContactItem oc, string emailAddressType, string emailAddress)
        {
            switch (emailAddressType)
            {
                case "EX":  // Microsoft Exchange address: "/o=xxxx/ou=xxxx/cn=Recipients/cn=xxxx"

                    Outlook.Application app = null;
                    Outlook.NameSpace ns = null;

                    try
                    {
                        app = oc.Application;
                        ns = app.GetNamespace("mapi");
                        // The emailEntryID is garbage (bug in Outlook 2007 and before?) - so we cannot do GetAddressEntryFromID().
                        // Instead we create a temporary recipient and ask Exchange to resolve it, then get the SMTP address from it.

                        Outlook.Recipient recipient = null;
                        try
                        {
                            recipient = ns.CreateRecipient(emailAddress);

                            recipient.Resolve();
                            if (recipient.Resolved)
                            {
                                Outlook.AddressEntry addressEntry = null;
                                try
                                {
                                    addressEntry = recipient.AddressEntry;
                                    if (addressEntry != null)
                                    {
                                        Outlook.ExchangeUser exchangeUser = null;
                                        try
                                        {
                                            exchangeUser = addressEntry.GetExchangeUser();
                                            if (exchangeUser != null)
                                            {
                                                return exchangeUser.PrimarySmtpAddress;
                                            }
                                            else
                                            {
                                                Log.Debug($"Error getting the email address of outlook contact '{oc.ToLogString()}' from Exchange format '{emailAddress}' and AddressEntryUserType '{addressEntry.AddressEntryUserType}'");
                                            }
                                        }
                                        finally
                                        {
                                            if (exchangeUser != null)
                                            {
                                                Marshal.ReleaseComObject(exchangeUser);
                                            }
                                        }
                                    }
                                }
                                finally
                                {
                                    if (addressEntry != null)
                                    {
                                        Marshal.ReleaseComObject(addressEntry);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (recipient != null)
                            {
                                Marshal.ReleaseComObject(recipient);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Warning($"Error getting the email address of outlook contact '{oc.ToLogString()}' from Exchange format '{emailAddress}': {ex.Message}");
                        return emailAddress;
                    }
                    finally
                    {
                        if (ns != null)
                        {
                            Marshal.ReleaseComObject(ns);
                        }
                        if (app != null)
                        {
                            Marshal.ReleaseComObject(app);
                        }
                    }

                    return emailAddress;

                case "SMTP":
                default:
                    return emailAddress;
            }
        }        
    }
}
