using Google.Apis.PeopleService.v1.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal static class ContactContentFingerprint
    {
        internal static string Capture(Outlook.ContactItem contact, bool useFileAs, bool syncPhotos)
        {
            var modified = contact.LastModificationTime;
            // Reuse the actual outbound field mappings, not an independent list of Outlook fields.
            var projected = new Person();
            ContactSync.UpdateContact(contact, projected, useFileAs);
            string photoHash = null;
            if (syncPhotos && contact.HasPhoto())
            {
                using (var photo = contact.GetOutlookPhoto())
                using (var stream = new MemoryStream())
                {
                    if (photo == null)
                        throw new ApplicationException("Cannot read Outlook contact photo for content tracking.");
                    photo.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                    photoHash = Hash(stream.ToArray());
                }
            }
            var fingerprint = Create(projected, Utilities.GetOutlookGroups(contact.Categories), useFileAs, syncPhotos, photoHash);
            if (modified != contact.LastModificationTime)
                throw new ApplicationException("Outlook contact changed while being read; retry the sync.");
            return fingerprint;
        }

        internal static string Create(Person projected, string[] categories, bool useFileAs, bool syncPhotos, string photoHash)
        {
            var fields = JObject.FromObject(projected, JsonSerializer.Create(new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            }));
            foreach (var field in new[] { "metadata", "etag", "resourceName", "clientData", "memberships", "photos" })
                fields.Remove(field);
            var snapshot = new JObject
            {
                ["fields"] = fields,
                ["categories"] = new JArray((categories ?? new string[0])
                    .Where(c => c != null).Select(c => c.Trim()).Where(c => c.Length > 0)
                    .Select(c => string.Equals(c, "Starred in Android", StringComparison.OrdinalIgnoreCase) ? "starred" : c)
                    .Distinct(StringComparer.Ordinal).OrderBy(c => c, StringComparer.Ordinal)),
                ["useFileAs"] = useFileAs,
                ["syncPhotos"] = syncPhotos,
                ["photo"] = syncPhotos ? photoHash : null
            };
            return "v1:" + Hash(Encoding.UTF8.GetBytes(Canonicalize(snapshot).ToString(Formatting.None)));
        }

        // Sort object properties and omit metadata; array ordering comes from the outbound mapper.
        private static JToken Canonicalize(JToken token)
        {
            if (token is JObject obj)
                return new JObject(obj.Properties().Where(p => p.Name != "metadata")
                    .OrderBy(p => p.Name, StringComparer.Ordinal)
                    .Select(p => new JProperty(p.Name, Canonicalize(p.Value))));
            if (token is JArray array)
                return new JArray(array.Select(Canonicalize));
            if (token.Type == JTokenType.String)
                return (string)token == null ? JValue.CreateNull()
                    : new JValue(((string)token).Replace("\r\n", "\n").Replace("\r", "\n"));
            return token.DeepClone();
        }

        internal static string Hash(byte[] value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(value)).Replace("-", "");
        }

        internal static string Key(params string[] parts)
        {
            return Hash(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(parts)));
        }
    }
}
