using Google.Apis.PeopleService.v1.Data;
using Microsoft.Office.Interop.Outlook;
using System;

namespace GoContactSyncMod
{
    /// <summary>
    /// Holds information about an Outlook contact during processing.
    /// We can not always instantiate an unlimited number of Exchange Outlook objects (policy limitations), 
    /// so instead we copy the info we need for our processing into instances of OutlookContactInfo and only
    /// get the real Outlook.ContactItem objects when needed to communicate with Outlook.
    /// </summary>
    internal class OutlookContactInfo
    {
        #region Internal classes
        internal class UserPropertiesHolder
        {
            public string GoogleContactId;
            public DateTime? LastSync;
            public string LastEtag;
        }
        #endregion

        #region Properties
        public string EntryID { get; set; }
        public string FileAs { get; set; }
        public string FullName { get; set; }
        public string TitleFirstLastAndSuffix { get; set; } //Additional unique identifier
        public string Email1Address { get; set; }
        public string MobileTelephoneNumber { get; set; }
        public string Categories { get; set; }
        public string Company { get; set; }
        public DateTime LastModificationTime { get; set; }
        public UserPropertiesHolder UserProperties { get; set; }
        #endregion

        #region Construction
        private OutlookContactInfo()
        {
            // Not public - we are always constructed from an Outlook.ContactItem (constructor below)
        }

        public OutlookContactInfo(ContactItem oc, Synchronizer sync)
        {
            UserProperties = new UserPropertiesHolder();
            Update(oc, sync);
        }
        #endregion

        public override string ToString()
        {
            var name = FileAs;
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name.RemoveNewLines();
            }

            name = FullName;
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name.RemoveNewLines();
            }

            name = TitleFirstLastAndSuffix;
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name.RemoveNewLines();
            }

            name = Company;
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name.RemoveNewLines();
            }

            name = Email1Address;
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name.RemoveNewLines();
            }

            return string.Empty;
        }

        internal void Update(ContactItem oc, Synchronizer sync)
        {
            EntryID = oc.EntryID;
            FileAs = oc.FileAs;
            FullName = oc.FullName;
            Email1Address = ContactPropertiesUtils.GetOutlookEmailAddress1(oc);
            MobileTelephoneNumber = oc.MobileTelephoneNumber;
            Categories = oc.Categories;

            //some contacts can throw "Not a legal OleAut date." exception when accessing LastModificationTime
            try
            {
                LastModificationTime = oc.LastModificationTime;
            }
            catch
            {
                LastModificationTime = DateTime.MinValue;
            }

            Company = oc.CompanyName;
            TitleFirstLastAndSuffix = GetTitleFirstLastAndSuffix(oc);
            UserProperties.GoogleContactId = ContactPropertiesUtils.GetOutlookGoogleId(oc);
            UserProperties.LastSync = ContactPropertiesUtils.GetOutlookLastSync(oc);
            UserProperties.LastEtag = ContactPropertiesUtils.GetOutlookLastEtag(oc);
        }

        internal ContactItem GetOriginalItemFromOutlook()
        {
            if (EntryID == null)
            {
                throw new ApplicationException("OutlookContactInfo cannot re-create the ContactItem from Outlook because EntryID is null, suggesting that this OutlookContactInfo was not created from an existing Outook contact.");
            }
            //"is" operator creates an implicit variable (COM leak), so unfortunately we need to avoid pattern matching
            var o = Synchronizer.OutlookNameSpace.GetItemFromID(EntryID);
#pragma warning disable IDE0019 // Use pattern matching
            var oc = o as ContactItem;
#pragma warning restore IDE0019 // Use pattern matching

            if (oc == null)
            {
                throw new ApplicationException("OutlookContactInfo cannot re-create the ContactItem from Outlook because there is no Outlook entry with this EntryID, suggesting that the existing Outook contact may have been deleted.");
            }

            return oc;
        }

        internal static string GetTitleFirstLastAndSuffix(ContactItem oc)
        {
            return ContactPropertiesUtils.GetTitleFirstLastAndSuffix(oc.Title, oc.FirstName, oc.MiddleName, oc.LastName, oc.Suffix);
        }
        


    }
}
