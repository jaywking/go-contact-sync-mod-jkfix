
using Serilog;
using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    class OutlookPropertiesUtils
    {
        public static string GetKey()
        {
            return "gos:oid:" + Synchronizer.SyncProfile;
        }

        internal static bool IsSameEmail(string e1, string e2)
        {
            var p1 = e1.ToLowerInvariant().Trim().Replace("@googlemail.", "@gmail.");
            var p2 = e2.ToLowerInvariant().Trim().Replace("@googlemail.", "@gmail.");

            if (p1.Equals(p2))
            {
                return true;
            }

            if (!p1.EndsWith("@gmail.com"))
            {
                return false;
            }

            if (!p2.EndsWith("@gmail.com"))
            {
                return false;
            }

            p1 = p1.Substring(0, p1.Length - 10).Replace(".", "");
            p2 = p2.Substring(0, p2.Length - 10).Replace(".", "");

            if (p1.Equals(p2))
            {
                return true;
            }

            var i1 = p1.IndexOf('+');
            var i2 = p2.IndexOf('+');

            if ((i1 < 0) && (i2 < 0))
            {
                return false;
            }

            if (i1 > 0)
            {
                p1 = p1.Substring(0, i1);
            }

            if (i2 > 0)
            {
                p2 = p2.Substring(0, i2);
            }

            if (p1.Equals(p2))
            {
                return true;
            }

            return false;
        }

        public static Outlook.UserProperty FindAndUnifySimilarProperty(Outlook.UserProperties up, string name, object fmt)
        {
            Outlook.UserProperty ret = null;

            var found = false;
            var t = Outlook.OlUserPropertyType.olText;
            dynamic v = null;

            var prefix_length = Synchronizer.OutlookUserPropertyPrefixTemplate.Length;
            var name_email = name.Substring(prefix_length, name.Length - prefix_length - 3);
            var n1 = name.Substring(name.Length - 3);
            var n2 = name.Substring(0, prefix_length);

            //TODO (obelix30) remove count use
            //[5/14/2020 9:53:17 AM | Debug] Exception: The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)
            //[5/14/2020 9:53:17 AM | Debug] Source: GOContactSync
            //[5 / 14 / 2020 9:53:17 AM | Debug] Stack Trace:    at Microsoft.Office.Interop.Outlook.UserProperties.get_Count()
            //at GoContactSyncMod.OutlookPropertiesUtils.FindAndUnifySimilarProperty(UserProperties up, String name) in .\OutlookPropertiesUtils.cs:line 74

            for (var i = up.Count; i > 0; i--)
            {
                Outlook.UserProperty p = null;

                try
                {
                    p = up[i];

                    if (p != null && p.Name != null && p.Name.Length - prefix_length - 3 > 0)
                    {
                        if (
                            string.Equals(p.Name.Substring(p.Name.Length - 3), n1, StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(p.Name.Substring(0, prefix_length), n2, StringComparison.OrdinalIgnoreCase) &&
                            IsSameEmail(p.Name.Substring(prefix_length, p.Name.Length - prefix_length - 3), name_email))
                        {
                            if (p.Name.Substring(prefix_length, p.Name.Length - prefix_length - 3) != name_email)
                            {
                                if (!found)
                                {
                                    found = true;
                                    t = p.Type;
                                    v = p.Value;
                                }

                                up.Remove(i);
                            }
                        }
                    }
                }
                finally
                {
                    if (p != null)
                    {
                        Marshal.ReleaseComObject(p);
                    }
                }
            }

            if (found)
            {
                ret = up.Add(name, t, false, fmt);
                ret.Value = v;

                return ret;
            }

            return ret;
        }

        internal static void RemoveOutlookPropertyIgnoreCase(Outlook.UserProperties up, string name)
        {
            var prefix_length = Synchronizer.OutlookUserPropertyPrefixTemplate.Length;
            var name_email = name.Substring(prefix_length, name.Length - prefix_length - 3);
            var n1 = name.Substring(name.Length - 3);
            var n2 = name.Substring(0, prefix_length);

            for (var i = up.Count; i > 0; i--)
            {
                Outlook.UserProperty p = null;

                try
                {
                    p = up[i];

                    if (p != null && p.Name != null && p.Name.Length - prefix_length - 3 > 0)
                    {
                        if (
                            string.Equals(p.Name.Substring(p.Name.Length - 3), n1, StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(p.Name.Substring(0, prefix_length), n2, StringComparison.OrdinalIgnoreCase) &&
                            IsSameEmail(p.Name.Substring(prefix_length, p.Name.Length - prefix_length - 3), name_email)
                            )
                        {
                            up.Remove(i);
                        }
                    }
                }
                finally
                {
                    if (p != null)
                    {
                        Marshal.ReleaseComObject(p);
                    }
                }
            }
        }

        public static bool SetOutlookGoogleId(Outlook.UserProperties userProps, string gid, string etag)
        {
            var changed = false;
            Outlook.UserProperty p1 = null;
            Outlook.UserProperty p2 = null;
            Outlook.UserProperty p3 = null;
            try
            {
                p1 = userProps[Synchronizer.OutlookPropertyNameId];

                //check if outlook entry already has google id property.
                if (p1 == null)
                {
                    //accessing user properties by using [] is case sensitive, but later calling up.Add fails to add new property 
                    //as it is case insensitive.  As workaround first remove all properties that are equal if we ignore the case
                    RemoveOutlookPropertyIgnoreCase(userProps, Synchronizer.OutlookPropertyNameId);

                    p1 = userProps.Add(Synchronizer.OutlookPropertyNameId, Outlook.OlUserPropertyType.olText, false, Outlook.OlFormatText.olFormatTextText);
                    changed = true;
                }

                if (p1.Value != gid)
                {
                    p1.Value = gid;
                    changed = true;
                }

                p2 = userProps[Synchronizer.OutlookPropertyNameSynced];
                if (p2 == null)
                {
                    //accessing user properties by using [] is case sensitive, but later calling up.Add fails to add new propery 
                    //as it is case insensitive.  As workaround first remove all properties that are equal if we ignore the case
                    RemoveOutlookPropertyIgnoreCase(userProps, Synchronizer.OutlookPropertyNameSynced);

                    p2 = userProps.Add(Synchronizer.OutlookPropertyNameSynced, Outlook.OlUserPropertyType.olDateTime, false, Outlook.OlFormatDateTime.olFormatDateTimeBestFit);
                    changed = true;
                }
                var now = DateTime.Now;
                p2.Value = now.AddSeconds(-now.Second);


                p3 = userProps[Synchronizer.OutlookPropertyNameEtag];

                //check if outlook entry already has google id property.
                if (p3 == null)
                {
                    //accessing user properties by using [] is case sensitive, but later calling up.Add fails to add new property 
                    //as it is case insensitive.  As workaround first remove all properties that are equal if we ignore the case
                    RemoveOutlookPropertyIgnoreCase(userProps, Synchronizer.OutlookPropertyNameEtag);

                    p3 = userProps.Add(Synchronizer.OutlookPropertyNameEtag, Outlook.OlUserPropertyType.olText, false, Outlook.OlFormatText.olFormatTextText);
                    changed = true;
                }

                if (p3.Value != etag)
                {
                    p3.Value = etag;
                    changed = true;
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                Log.Debug("Name: " + Synchronizer.OutlookPropertyNameId);
                Log.Debug("Value: " + gid);
                Log.Debug("Name: " + Synchronizer.OutlookPropertyNameSynced);
                Log.Debug("Value: " + DateTime.Now);
                Log.Debug("Name: " + Synchronizer.OutlookPropertyNameEtag);
                Log.Debug("Value: " + etag);
                throw;
            }
            finally
            {
                if (p1 != null)
                {
                    Marshal.ReleaseComObject(p1);
                }

                if (p2 != null)
                {
                    Marshal.ReleaseComObject(p2);
                }

                if (p3 != null)
                {
                    Marshal.ReleaseComObject(p3);
                }
            }

            return changed;
        }

        public static DateTime? GetOutlookLastSync(Outlook.UserProperties userProps)
        {
            Outlook.UserProperty p = null;
            try
            {
                p = userProps[Synchronizer.OutlookPropertyNameSynced];

                if (p == null)
                {
                    //accessing user properties by using [] is case sensitive, but later calling up.Add fails to add new property 
                    //as it is case insensitive.  As workaround first remove all properties that are equal if we ignore the case
                    p = FindAndUnifySimilarProperty(userProps, Synchronizer.OutlookPropertyNameSynced, Outlook.OlFormatDateTime.olFormatDateTimeBestFit);
                }

                if (p != null)
                {
                    var s = Convert.ToString(p.Value);

                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        if (DateTime.TryParse(s, out DateTime lastSync))
                        {
                            if (lastSync == null)
                                lastSync = DateTime.MinValue;

                            if (lastSync.Kind == DateTimeKind.Utc)
                                lastSync = TimeZoneInfo.ConvertTimeFromUtc(lastSync, TimeZoneInfo.Local);
                            return lastSync;
                        }
                    }
                }
            }
            finally
            {
                if (p != null)
                {
                    Marshal.ReleaseComObject(p);
                }

            }
            return (DateTime?)null;
        }

        public static string GetOutlookLastEtag(Outlook.UserProperties userProps)
        {
            Outlook.UserProperty p = null;
            try
            {
                p = userProps[Synchronizer.OutlookPropertyNameEtag];

                if (p == null)
                {
                    //accessing user properties by using [] is case sensitive, but later calling up.Add fails to add new property 
                    //as it is case insensitive.  As workaround first remove all properties that are equal if we ignore the case
                    p = FindAndUnifySimilarProperty(userProps, Synchronizer.OutlookPropertyNameEtag, Outlook.OlFormatText.olFormatTextText);
                }

                if (p != null)
                {
                    var s = Convert.ToString(p.Value);

                    return s;
                }
            }
            finally
            {
                if (p != null)
                {
                    Marshal.ReleaseComObject(p);
                }

            }
            return string.Empty;
        }

        public static T GetOutlookPropertyValue<T>(string key, Outlook.UserProperties up, object fmt)
        {
            T val = default;

            Outlook.UserProperty p = null;
            try
            {

                p = up[key];

                if (p == null)
                {
                    //accessing user properties by using [] is case sensitive, but later calling up.Add fails to add new property 
                    //as it is case insensitive.  As workaround first remove all properties that are equal if we ignore the case
                    p = FindAndUnifySimilarProperty(up, key, fmt);
                }

                if (p != null)
                {
                    val = (T)p.Value;
                }
            }
            finally
            {
                if (p != null)
                {
                    Marshal.ReleaseComObject(p);
                }

            }
            return val;
        }


        public static void ResetOutlookGoogleId(Outlook.UserProperties up)
        {
            if (up != null)
                for (var i = up.Count; i > 0; i--)
                {
                    Outlook.UserProperty p = null;

                    try
                    {
                        p = up[i];
                        if (p.Name == Synchronizer.OutlookPropertyNameId || p.Name == Synchronizer.OutlookPropertyNameSynced || p.Name == Synchronizer.OutlookPropertyNameEtag)
                        {
                            up.Remove(i);
                        }
                    }
                    finally
                    {
                        if (p != null)
                        {
                            Marshal.ReleaseComObject(p);
                        }
                    }
                }
        }
    }
}
