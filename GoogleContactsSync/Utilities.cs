using Dangl.TextConverter.Rtf;
using Google.Apis.PeopleService.v1.Data;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Net;
using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace GoContactSyncMod
{
    internal static class Utilities
    {
        public static string GetTempFileName(string ext)
        {
            var fileName = Path.GetRandomFileName();
            fileName = Path.ChangeExtension(fileName, ext);
            fileName = Path.Combine(Path.GetTempPath(), fileName);
            return fileName;
        }

        public static byte[] BitmapToBytes(Bitmap bmp)
        {
            //bitmap
            using (var stream = new MemoryStream())
            {
                bmp.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);
                return stream.ToArray();
            }
        }

        public static bool HasContactPhoto(Person gc)
        {
            if (gc.Photos == null)
            {
                Log.Debug($"Google Contact has no photo!");
                return false;
            }

            foreach (var photo in gc.Photos)
            {
                if (photo != null && photo.Metadata != null && photo.Metadata.Source != null && photo.Metadata.Source.Type != null && photo.Metadata.Source.Type.Trim().Equals(ContactPropertiesUtils.CONTACT, StringComparison.OrdinalIgnoreCase))
                {
                    var isDefaultPhoto = photo.Default__ ?? false;

                    if (!isDefaultPhoto) //don'T consider a default avatar photo as real Google contact photo
                    {
                        Log.Debug($"Google Contact has a contact photo!");
                        return true;
                    }
                }

                //else: photo is not a contact photo (e.g. profile photo), skipping.                
            }
            Log.Debug($"Google Contact has no contact photo!");
            return false;
        }

            //public static bool SaveGooglePhoto(Synchronizer sync, Person gc, Image image)
            //{
            //    if (gc.ContactEntry.PhotoUri == null)
            //    {
            //        throw new Exception("Must reload contact from google.");
            //    }

            //    try
            //    {
            //        using (var client = new WebClient())
            //        {
            //            //client.Headers.Add(HttpRequestHeader.Authorization, "Bearer " + sync.PeopleRequest.Settings.OAuth2Parameters.AccessToken);//ToDo: Check, how to get the AccessToken from new PeopleAPI
            //            client.Headers.Add(HttpRequestHeader.ContentType, "image/*");
            //            using (var pic = new Bitmap(image))
            //            {
            //                using (var s = client.OpenWrite(gc.ContactEntry.PhotoUri.AbsoluteUri, "PUT"))
            //                {
            //                    var bytes = BitmapToBytes(pic);
            //                    s.Write(bytes, 0, bytes.Length);
            //                    s.Flush();
            //                }
            //            }
            //        }
            //    }
            //    catch
            //    {
            //        return false;
            //    }
            //    return true;
            //}

        public static Image GetGoogleContactPhoto(Synchronizer sync, Person gc)
        {
            try
            {
                Photo photo = null;
                if (gc.Photos != null)
                {
                    foreach (var ph in gc.Photos)
                    {
                        if (ph != null && ph.Metadata != null && ph.Metadata.Source != null && ph.Metadata.Source.Type != null && ph.Metadata.Source.Type.Equals(ContactPropertiesUtils.CONTACT, StringComparison.OrdinalIgnoreCase))
                        {
                            var isDefaultPhoto = ph.Default__ ?? false;

                            if (!isDefaultPhoto)
                            {
                                photo = ph;
                                break;
                            }                            
                        }
                        //else: photo is not a contact photo (e.g. profile photo), try next.
                    }
                }
                if (photo != null)
                {
                    using (var client = new WebClient())
                    {
                        client.Headers.Add(HttpRequestHeader.Authorization, "Bearer " + sync.GooglePeopleRequest.AccessToken); //ToDo: Check, if works from new PeopleAPI, seems to work (big-r, 08.07.2021)!
                        var stream = client.OpenRead(photo.Url);
                        var reader = new BinaryReader(stream);
                        var image = Image.FromStream(stream);
                        reader.Close();
                        return image;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, $"Error fetching Google contact photo: {ex}");
            }
            return null;
        }


        public static Image CropImageGoogleFormat(Image original)
        {
            // crop image to a square in the center
            if (original.Height == original.Width)
            {
                return original;
            }

            if (original.Height > original.Width)
            {
                // tall image
                var width = original.Width;
                var height = width;

                var diff = original.Height - height;
                var p = new Point(0, diff / 2);
                var r = new Rectangle(p, new Size(width, height));

                return CropImage(original, r);
            }
            else
            {
                // flat image
                var height = original.Height;
                var width = height;

                var diff = original.Width - width;
                var p = new Point(diff / 2, 0);
                var r = new Rectangle(p, new Size(width, height));

                return CropImage(original, r);
            }
        }

        public static Image CropImage(Image original, Rectangle cropArea)
        {
            using (var bmpImage = new Bitmap(original))
            {
                var bmpCrop = bmpImage.Clone(cropArea, bmpImage.PixelFormat);
                return bmpCrop;
            }
        }

        public static bool ContainsGroup(Synchronizer sync, Person gc, string groupName)
        {
            var group = sync.GetGoogleGroupByName(groupName);
            return group != null && ContainsGroup(gc, group);
        }

        public static bool ContainsGroup(Person gc, ContactGroup group)
        {
            if (gc.Memberships != null)
                foreach (var m in gc.Memberships)
                {
                    if (m.ContactGroupMembership != null && m.ContactGroupMembership.ContactGroupResourceName == group.ResourceName)
                    {
                        return true;
                    }
                }
            return false;
        }

        public static bool ContainsGroup(Outlook.ContactItem oc, string group)
        {
            return oc.Categories != null && oc.Categories.Contains(group);
        }

        public static Collection<ContactGroup> GetGoogleGroups(Synchronizer sync, Person gc)
        {
            var groups = new Collection<ContactGroup>();

            if (gc.Memberships != null && gc.Memberships.Count > 0)
                foreach (var group in gc.Memberships)
                    if (group != null && group.ContactGroupMembership != null && group.ContactGroupMembership.ContactGroupResourceName != Synchronizer.myContactsGroup)
                    {
                        var g = sync.GetGoogleGroupByResourceName(group.ContactGroupMembership.ContactGroupResourceName);
                        if (g != null)
                            groups.Add(g);
                    }
            return groups;
        }

        public static void AddGoogleGroup(Person gc, ContactGroup group)
        {
            if (ContainsGroup(gc, group))
            {
                return;
            }

            var m = new Membership()
            {
                ContactGroupMembership = new ContactGroupMembership()
                {
                    ContactGroupResourceName = group.ResourceName
                }
                //HRef = group.GroupEntry.Id.AbsoluteUri;
            };
            if (gc.Memberships == null)
                gc.Memberships = new List<Membership>();
            gc.Memberships.Add(m);
        }

        public static void RemoveGoogleGroup(Person gc, ContactGroup group)
        {
            if (!ContainsGroup(gc, group))
            {
                return;
            }

            // TODO: broken. removes group membership but does not remove contact
            // from group in the end.

            // look for id
            Membership mem;
            for (var i = 0; i < gc.Memberships.Count; i++)
            {
                mem = gc.Memberships[i];
                if (mem.ContactGroupMembership != null && mem.ContactGroupMembership.ContactGroupResourceName == group.ResourceName) //Todo: Check
                {
                    gc.Memberships.Remove(mem);
                    return;
                }
            }
            throw new Exception("Did not find group");
        }

        public static string[] GetOutlookGroups(string outlookContactCategories)
        {
            if (outlookContactCategories == null)
            {
                return new string[] { };
            }

            var listseparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator.ToCharArray();
            if (!outlookContactCategories.Contains(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator))
            {// ListSeparator doesn't work always, because ListSeparator returns "," instead of ";"
                listseparator = ",".ToCharArray();
                if (!outlookContactCategories.Contains(","))
                {
                    listseparator = ";".ToCharArray();
                }
            }
            var categories = outlookContactCategories.Split(listseparator);

            for (var i = 0; i < categories.Length; i++)
            {
                categories[i] = categories[i].Trim();
            }
            return categories;
        }

        public static void AddOutlookGroup(Outlook.ContactItem oc, string group)
        {
            if (ContainsGroup(oc, group))
            {
                return;
            }

            // append
            if (oc.Categories == null)
            {
                oc.Categories = "";
            }

            if (!string.IsNullOrEmpty(oc.Categories))
            {
                oc.Categories += ", " + group;
            }
            else
            {
                oc.Categories += group;
            }
        }

        public static void RemoveOutlookGroup(Outlook.ContactItem oc, string group)
        {
            if (!ContainsGroup(oc, group))
            {
                return;
            }
            oc.Categories = oc.Categories.Replace(", " + group, "");
            oc.Categories = oc.Categories.Replace(group, "");
        }

        public static string ConvertToText(string rtf)
        {
            return RtfToText.ConvertRtfToText(rtf);
        }

        public static string ConvertToText(byte[] rtf)
        {
            if (rtf != null)
            {
                var encoding = new System.Text.ASCIIEncoding();
                return ConvertToText(encoding.GetString(rtf));
            }
            return string.Empty;
        }

        /// <summary>
        /// Time tolerance in seconds - used when comparing date modified.
        /// Less than 60 seconds doesn't make sense, as the lastSync is saved without seconds and if it is compared
        /// with the LastUpdate dates of Google and Outlook, in the worst case you compare e.g. 15:59 with 16:00 and 
        /// after truncating to minutes you compare 15:00 wiht 16:00
        /// Better take 120 seconds, because when resetting matches the time difference can be up to 2 minutes
        /// </summary>
        private static int TimeTolerance = 120;
        internal static bool UpdatedSinceLastSync(DateTime d, DateTime s)
        {
            return (int)d.Subtract(s).TotalSeconds > TimeTolerance;
        }
    }

    public class OutlookFolder : IComparable
    {
        public OutlookFolder(string folderName, string folderID, bool isDefaultFolder)
        {
            FolderName = folderName;
            FolderID = folderID;
            IsDefaultFolder = isDefaultFolder;
        }

        public string FolderName { get; }

        public string FolderID { get; }

        public bool IsDefaultFolder { get; }

        public string DisplayName => FolderName + (IsDefaultFolder ? " (Default)" : string.Empty);

        public int CompareTo(object obj)
        {
            if (obj == null)
            {
                return 1;
            }

            var other = obj as OutlookFolder;
            if (other == null)
            {
                throw new ArgumentException($"Cannot compare {GetType()} with {obj.GetType()}");
            }
            return CompareTo(this, other);
        }

        public static bool operator <(OutlookFolder left, OutlookFolder right)
        {
            if (left is null)
            {
                return !(right is null);
            }
            else if (right is null)
            {
                return false;
            }

            return CompareTo(left, right) < 0;
        }

        public static bool operator >(OutlookFolder left, OutlookFolder right)
        {
            if (left is null)
            {
                return right is null;
            }
            else if (right is null)
            {
                return true;
            }

            return CompareTo(left, right) > 0;
        }

        public static bool operator ==(OutlookFolder left, OutlookFolder right)
        {
            if (left is null)
            {
                return right is null;
            }
            else if (right is null)
            {
                return false;
            }

            return Equals(left, right);
        }

        public static bool operator !=(OutlookFolder left, OutlookFolder right)
        {
            return !(left == right);
        }

        public override bool Equals(object obj)
        {
            return !(obj is null) && obj.GetType() == GetType() && Equals(this, obj as OutlookFolder);
        }

        internal static bool Equals(OutlookFolder left, OutlookFolder right)
        {
            return (right.FolderName == left.FolderName) &&
                  (right.FolderID == left.FolderID) &&
                  (right.IsDefaultFolder == left.IsDefaultFolder);
        }

        internal static int CompareTo(OutlookFolder left, OutlookFolder right)
        {
            var _folderNameComparison = left.FolderName.CompareTo(right.FolderName);
            if (_folderNameComparison != 0)
            {
                return _folderNameComparison;
            }

            var _folderIDComparison = left.FolderID.CompareTo(right.FolderID);
            return _folderIDComparison != 0 ? _folderIDComparison : left.IsDefaultFolder.CompareTo(right.IsDefaultFolder);
        }

        public override int GetHashCode()
        {
            return HashUtils.CombineHashCodes(FolderName.GetHashCode(), FolderID.GetHashCode(), IsDefaultFolder.GetHashCode());
        }

        public override string ToString()
        {
            return this is OutlookFolder ? FolderID : base.ToString();
        }
    }

    public class GoogleCalendar : IComparable
    {
        public GoogleCalendar(string folderName, string folderID, bool isDefaultFolder)
        {
            FolderName = folderName;
            FolderID = folderID;
            IsDefaultFolder = isDefaultFolder;
        }

        public string FolderName { get; }

        public string FolderID { get; }

        public bool IsDefaultFolder { get; }

        public string DisplayName => FolderName + (IsDefaultFolder ? " (Default)" : string.Empty);

        public int CompareTo(object obj)
        {
            if (obj == null)
            {
                return 1;
            }

            var other = obj as GoogleCalendar;
            if (other == null)
            {
                throw new ArgumentException($"Cannot compare {GetType()} with {obj.GetType()}");
            }
            return CompareTo(this, other);
        }

        public static bool operator <(GoogleCalendar left, GoogleCalendar right)
        {
            if (left is null)
            {
                return !(right is null);
            }
            else if (right is null)
            {
                return false;
            }

            return CompareTo(left, right) < 0;
        }

        public static bool operator >(GoogleCalendar left, GoogleCalendar right)
        {
            if (left is null)
            {
                return right is null && false;
            }
            else if (right is null)
            {
                return true;
            }

            return CompareTo(left, right) > 0;
        }

        public static bool operator ==(GoogleCalendar left, GoogleCalendar right)
        {
            if (left is null)
            {
                return right is null;
            }
            else if (right is null)
            {
                return false;
            }

            return Equals(left, right);
        }

        public static bool operator !=(GoogleCalendar left, GoogleCalendar right)
        {
            return !(left == right);
        }

        public override bool Equals(object obj)
        {
            return !(obj is null) && obj.GetType() == GetType() && Equals(this, obj as OutlookFolder);
        }

        internal static bool Equals(GoogleCalendar left, GoogleCalendar right)
        {
            return (right.FolderName == left.FolderName) &&
                  (right.FolderID == left.FolderID) &&
                  (right.IsDefaultFolder == left.IsDefaultFolder);
        }

        internal static int CompareTo(GoogleCalendar left, GoogleCalendar right)
        {
            var _folderNameComparison = left.FolderName.CompareTo(right.FolderName);
            if (_folderNameComparison != 0)
            {
                return _folderNameComparison;
            }

            var _folderIDComparison = left.FolderID.CompareTo(right.FolderID);
            return _folderIDComparison != 0 ? _folderIDComparison : left.IsDefaultFolder.CompareTo(right.IsDefaultFolder);
        }

        public override int GetHashCode()
        {
            return HashUtils.CombineHashCodes(FolderName.GetHashCode(), FolderID.GetHashCode(), IsDefaultFolder.GetHashCode());
        }

        public override string ToString()
        {
            return this is GoogleCalendar ? FolderID : base.ToString();
        }

    }

    //Taken from Tuple
    public static class HashUtils
    {
        public static int CombineHashCodes(int h1, int h2)
        {
            return ((h1 << 5) + h1) ^ h2;
        }

        public static int CombineHashCodes(int h1, int h2, int h3)
        {
            return CombineHashCodes(CombineHashCodes(h1, h2), h3);
        }
    }

    
}
