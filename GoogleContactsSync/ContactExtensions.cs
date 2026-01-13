using Google.Apis.PeopleService.v1.Data;
using Serilog;

namespace GoContactSyncMod
{
    public static class ContactExtensions
    {
        public static void ToDebugLog(this Person c)
        {
            Log.Debug("*** Google contact ***");
            /*if (c.AppControl != null)
            {
                Log.Debug(" - AppControl:");
            }
            if (c.AtomEntry != null)
            {
                Log.Debug(" - AtomEntry:");
            }
            Log.Debug(" - Author: " + (c.Author ?? "null"));
            if (c.BatchData != null)
            {
                Log.Debug(" - BatchData:");
                if (c.BatchData.Id != null)
                {
                    Log.Debug("  - Id: " + c.BatchData.Id);
                }
            }
            if (c.Categories != null)
            {
                Log.Debug(" - Categories:");
                foreach (var e in c.Categories)
                {
                    Log.Debug("  - Scheme: " + (e.Scheme ?? "null"));
                    Log.Debug("  - Term: " + (e.Term ?? "null"));
                }
            }
            if (c.ContactEntry != null)
            {
                Log.Debug(" - ContactEntry:");
                Log.Debug("  - Initials: " + (c.ContactEntry.Initials ?? "null"));
            }
            Log.Debug(" - Content: " + (c.Content ?? "null"));*/
            Log.Debug(" - Deleted: " + ((c.Metadata != null)?(c.Metadata.Deleted??false).ToString(): "null"));
            if (c.EmailAddresses != null)
            {
                Log.Debug(" - EmailAddresses:");
                foreach (var e in c.EmailAddresses)
                if (e != null)
                    {
                        Log.Debug("  - Address: " + (e.Value ?? "null"));
                        Log.Debug("  - Label: " + (e.Type ?? "null"));
                        //Log.Debug("  - Primary: " + e.Primary.ToString());
                    }
            }
            Log.Debug(" - ETag: " + (c.ETag ?? "null"));
            if (c.ClientData != null)
            {
                Log.Debug(" - ClientData:");
                foreach (var e in c.ClientData)
                    if (e != null)
                    {
                        Log.Debug("  - Name: " + (e.Key ?? "null"));
                        Log.Debug("  - Value: " + (e.Value ?? "null"));
                    }
            }
            if (c.Memberships != null)
            {
                Log.Debug(" - Membership:");
                foreach (var e in c.Memberships)
                    if (e != null)
                    {
                        Log.Debug("  - HRef: " + (e.ToString() ?? "null"));
                    }
            }
            Log.Debug(" - Id: " + (c.ResourceName ?? "null"));
            if (c.ImClients != null)
            {
                Log.Debug(" - ImClients:");
                foreach (var e in c.ImClients)
                    if (e != null)
                    {
                        Log.Debug("  - Value: " + (e.ToString() ?? "null"));
                    }
            }
            //Log.Debug(" - IsDraft: " + c.IsDraft.ToString());
            /*if (c.Languages != null)
            {
                Log.Debug(" - Languages:");
                foreach (var e in c.Languages)
                {
                    Log.Debug("  - Value: " + (e.Value ?? "null"));
                }
            }
            Log.Debug(" - Location: " + (c.Location ?? "null"));
            if (c.MediaSource != null)
            {
                Log.Debug(" - MediaSource:");
                Log.Debug("  - Name: " + (c.MediaSource.Name ?? "null"));
            }*/
            var name = ContactPropertiesUtils.GetGooglePrimaryName(c);            
            if (name != null)
            {
                Log.Debug(" - Name:");
                Log.Debug("  - FamilyName: " + (name.FamilyName ?? "null"));
                Log.Debug("  - UnstructuredName: " + (name.UnstructuredName ?? "null"));
                Log.Debug("  - GivenName: " + (name.GivenName ?? "null"));
            }
            if (c.Organizations != null)
            {
                Log.Debug(" - Organizations:");
                foreach (var e in c.Organizations)
                    if (e != null)
                    {
                        Log.Debug("  - Name: " + (e.Name ?? "null"));
                    }
            }
            if (c.PhoneNumbers != null)
            {
                Log.Debug(" - PhoneNumbers:");
                foreach (var e in c.PhoneNumbers)
                    if (e != null)
                    {
                        Log.Debug("  - Rel: " + (e.Type ?? "null"));
                        Log.Debug("  - Value: " + (e.Value ?? "null"));
                    }
            }
            Log.Debug(" - PhotoEtag: " + ((c.Photos != null && c.Photos.Count > 0)?(c.Photos[0].ETag ?? "null"):"null"));
            /*if (c.PhotoUri != null)
            {
                Log.Debug(" - PhotoUri:");
                Log.Debug("  - OriginalString: " + (c.PhotoUri.OriginalString ?? "null"));
            }*/
            if (c.Addresses != null)
            {
                Log.Debug(" - Addresses:");
                foreach (var e in c.Addresses)
                    if (e != null)
                    {
                        Log.Debug("  - Street: " + (e.StreetAddress ?? "null"));
                    }
            }
            /*if (c.PrimaryEmail != null)
            {
                Log.Debug(" - PrimaryEmail:");
                Log.Debug("  - Value: " + (c.PrimaryEmail.Value ?? "null"));
            }
            if (c.PrimaryIMAddress != null)
            {
                Log.Debug(" - PrimaryIMAddress:");
                Log.Debug("  - Value: " + (c.PrimaryIMAddress.Value ?? "null"));
            }
            if (c.PrimaryPhonenumber != null)
            {
                Log.Debug(" - PrimaryPhonenumber:");
                Log.Debug("  - Value: " + (c.PrimaryPhonenumber.Value ?? "null"));
            }
            if (c.PrimaryPostalAddress != null)
            {
                Log.Debug(" - PrimaryPostalAddress:");
                Log.Debug("  - Street: " + (c.PrimaryPostalAddress.StreetAddress ?? "null"));
            }
            Log.Debug(" - ReadOnly: " + c.ReadOnly.ToString());
            Log.Debug(" - Self: " + (c.Self ?? "null"));
            Log.Debug(" - Summary: " + (c.Summary ?? "null"));
            Log.Debug(" - Title: " + (c.Title ?? "null"));
            Log.Debug(" - Updated: " + (c.Updated != null ? c.Updated.ToString() : "null"));*/
            Log.Debug("*** Google contact ***");
        }

        public static string ToLogString(this Person gc)
        {

            var gn = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc);
            if (gn != null)
            {
                gn = gn.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();
                if (!string.IsNullOrWhiteSpace(gn))
                {
                    return gn;
                }
            }

            var name = ContactPropertiesUtils.GetGooglePrimaryName(gc);
            if (name != null)
                gn = name.UnstructuredName;
            if (gn != null)
            {
                gn = gn.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();
                if (!string.IsNullOrWhiteSpace(gn))
                {
                    return gn;
                }
            }

            gn = ContactPropertiesUtils.GetGoogleTitleFirstLastAndSuffix(gc);
            if (gn != null)
            {
                return gn.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();
            }

            if (gc.Organizations.Count > 0)
            {
                gn = gc.Organizations[0].Name;
                if (gn != null)
                {
                    return gn.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();
                }
            }

            return string.Empty;
        }
    }
}