using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    public static class ContactItemExtensions
    {
        public static bool HasPhoto(this Outlook.ContactItem oc)
        {
            return oc.HasPicture;
        }

        public static bool SetOutlookPhoto(this Outlook.ContactItem oc, string fullImagePath)
        {
            try
            {
                oc.AddPicture(fullImagePath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetTempFileName(string ext)
        {
            var fileName = Path.GetRandomFileName();
            fileName = Path.ChangeExtension(fileName, ext);
            fileName = Path.Combine(Path.GetTempPath(), fileName);
            return fileName;
        }

        public static bool SetOutlookPhoto(this Outlook.ContactItem oc, Image image)
        {
            var fn = GetTempFileName("jpg");
            try
            {
                image.Save(fn);
                return SetOutlookPhoto(oc, fn);
            }
            catch
            {
                return false;
            }
            finally
            {
                File.Delete(fn);
            }
        }
        public static Bitmap GetOutlookPhoto(this Outlook.ContactItem oc)
        {
            if (!HasPhoto(oc))
            {
                return null;
            }

            Outlook.Attachments attachments = null;
            try
            {
                attachments = oc.Attachments;

                if (attachments == null)
                {
                    return null;
                }

                for (var i = attachments.Count; i > 0; i--)
                {
                    Outlook.Attachment a = null;
                    try
                    {
                        a = attachments[i];
                        var s = a.DisplayName.ToUpper();

                        if (s.Contains("CONTACTPICTURE") || s.Contains("CONTACTPHOTO"))
                        {
                            var fn = GetTempFileName("jpg");
                            a.SaveAsFile(fn);

                            try
                            {
                                using (var fs = new FileStream(fn, FileMode.Open))
                                {
                                    using (var img = Image.FromStream(fs))
                                    {
                                        return new Bitmap(img);
                                    }
                                }
                            }
                            finally
                            {
                                File.Delete(fn);
                            }
                        }
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
            catch
            {
                return null;
            }
            finally
            {
                if (attachments != null)
                {
                    Marshal.ReleaseComObject(attachments);
                }
            }
            return null;
        }


        public static string ToLogString(this Outlook.ContactItem oc)
        {
            string name;
            try
            {
                name = oc.FileAs;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                }
            }
            catch
            {
                return "Exception while accessing FileAs";
            }

            try
            {
                name = oc.FullName;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                }
            }
            catch
            {
                return "Exception while accessing FullName";
            }

            try
            {
                name = oc.Email1Address;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                }
            }
            catch
            {
                return "Exception while accessing Email1Address";
            }

            return string.Empty;
        }
    }
}
