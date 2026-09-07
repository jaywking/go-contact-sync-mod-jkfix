using Newtonsoft.Json;
using System;
using System.IO;

namespace GoContactSyncMod
{
    internal enum ContactContentDecision { BaselineRecorded, Unchanged, Update }

    // One atomic file per contact/profile/account. No contact field values are persisted.
    internal sealed class ContactContentTracker
    {
        private readonly string directory;

        internal ContactContentTracker(string directory)
        {
            this.directory = directory;
        }

        private sealed class Baseline
        {
            public int Version { get; set; } = 1;
            public string Fingerprint { get; set; }
            public string GoogleLink { get; set; }
            public bool Pending { get; set; }
            public bool Verified { get; set; }
        }

        internal ContactContentDecision Check(string key, string googleResource, string fingerprint, bool legacyUpdateRequired)
        {
            var baseline = Read(key);
            if (baseline == null)
            {
                if (legacyUpdateRequired)
                    return ContactContentDecision.Update;
                // Deliberate migration policy: observe existing contents, without rewriting Google.
                // This is an unverified starting baseline, not a claim that Google has these values.
                Write(key, new Baseline { Fingerprint = fingerprint, GoogleLink = Link(googleResource) });
                return ContactContentDecision.BaselineRecorded;
            }
            return baseline.Pending || baseline.GoogleLink != Link(googleResource) || baseline.Fingerprint != fingerprint
                ? ContactContentDecision.Update : ContactContentDecision.Unchanged;
        }

        internal bool SaveUpdate(string key, string fingerprint, Func<string> saveContact)
        {
            var baseline = Read(key) ?? new Baseline();
            // Persist retry intent before any remote write. Even a first update or failed photo
            // save must retry after restart, regardless of changes to Outlook's sync timestamp.
            baseline.Pending = true;
            Write(key, baseline);
            var googleResource = saveContact();
            if (string.IsNullOrEmpty(googleResource))
                return false;
            // Use the snapshot captured before mapping/sending, never a fresh post-save snapshot:
            // an edit made during the request must still be detected on the next run.
            Write(key, new Baseline
            {
                Fingerprint = fingerprint, GoogleLink = Link(googleResource), Verified = true
            });
            return true;
        }

        private static string Link(string resource) => ContactContentFingerprint.Key(resource ?? "");

        private string PathFor(string key)
        {
            if (key == null || key.Length != 64 || key.Trim("0123456789ABCDEF".ToCharArray()).Length != 0)
                throw new ArgumentException("Invalid content tracking key.", nameof(key));
            return Path.Combine(directory, key + ".json");
        }

        private Baseline Read(string key)
        {
            var path = PathFor(key);
            if (!File.Exists(path))
                return null;
            // Do not silently re-baseline a corrupt cache and potentially swallow pending edits.
            var baseline = JsonConvert.DeserializeObject<Baseline>(File.ReadAllText(path));
            if (baseline == null || baseline.Version != 1 ||
                (!baseline.Pending && (string.IsNullOrEmpty(baseline.Fingerprint) || string.IsNullOrEmpty(baseline.GoogleLink))))
                throw new InvalidDataException("Invalid contact content baseline: " + path);
            return baseline;
        }

        private void Write(string key, Baseline baseline)
        {
            Directory.CreateDirectory(directory);
            var path = PathFor(key);
            var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temp, JsonConvert.SerializeObject(baseline));
                if (File.Exists(path))
                    File.Replace(temp, path, null);
                else
                    File.Move(temp, path);
            }
            finally
            {
                if (File.Exists(temp))
                    File.Delete(temp);
            }
        }
    }
}
