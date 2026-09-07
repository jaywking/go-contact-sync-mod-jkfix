using Google.Apis.PeopleService.v1.Data;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    [NonParallelizable]
    public class ContactPreparationIsolationTests
    {
        private static ContactMatch Contact(string name)
        {
            var outlook = (OutlookContactInfo)Activator.CreateInstance(typeof(OutlookContactInfo), true);
            outlook.FileAs = name;
            outlook.UserProperties = new OutlookContactInfo.UserPropertiesHolder { GoogleContactId = "test" };
            return new ContactMatch(outlook, new Person
            {
                ResourceName = "people/" + name,
                ClientData = new List<ClientData>
                {
                    new ClientData { Key = OutlookPropertiesUtils.GetKey(), Value = "outlook-" + name }
                }
            });
        }

        private static Synchronizer Sync(params ContactMatch[] contacts)
        {
            var sync = new Synchronizer();
            typeof(Synchronizer).GetProperty(nameof(Synchronizer.Contacts)).GetSetMethod(true)
                .Invoke(sync, new object[] { new List<ContactMatch>(contacts) });
            return sync;
        }

        [TestCase(false)]
        [TestCase(true)]
        public void PartialPreparationIsNeverSavedAndOtherContactsContinue(bool subscribe)
        {
            var before = Contact("before");
            var broken = Contact("broken");
            var after = Contact("after");
            var sync = Sync(before, broken, after);
            Exception reported = null;
            if (subscribe) sync.ErrorEncountered += (title, error) => reported = error;
            var failure = new IOException("Cannot read contact photo");
            ContactsMatcher.SyncContacts(sync, (match, _) =>
            {
                match.GoogleContactDirty = true;
                match.PendingContentFingerprint = "partly-prepared";
                if (match == broken) throw failure;
            });
            var saved = new List<ContactMatch>();
            sync.SaveContacts(sync.Contacts, match => sync.SaveContact(match, item =>
            {
                saved.Add(item);
                return true;
            }));
            Assert.That(saved, Is.EqualTo(new[] { before, after }));
            Assert.That(sync.ErrorCount, Is.EqualTo(1));
            Assert.That(sync.SyncedCount, Is.EqualTo(2));
            Assert.That(broken.PreparationFailed, Is.True);
            Assert.That(sync.SaveContact(broken, _ => throw new Exception("Unexpected write")), Is.False);
            if (subscribe)
            {
                Assert.That(reported.InnerException, Is.SameAs(failure));
                Assert.That(reported.Message, Does.Contain("broken").And.Contain(failure.Message));
            }
        }

        [Test]
        public void FailedPreparationCannotTurnIntoADeletion()
        {
            var broken = new ContactMatch(null, new Person { ResourceName = "people/broken" });
            var sync = Sync(broken);
            sync.SyncOption = SyncOption.OutlookToGoogleOnly;
            sync.SyncDelete = true;
            ContactsMatcher.SyncContacts(sync, (match, _) => throw new IOException("Unreadable baseline"));
            Assert.DoesNotThrow(() => sync.SaveContacts(sync.Contacts));
            Assert.That(sync.DeletedCount, Is.Zero);
            Assert.That(sync.ErrorCount, Is.EqualTo(1));
        }

        [Test]
        public void CorruptBaselineIsRetainedWhileOtherContactSyncs()
        {
            var directory = Path.Combine(Path.GetTempPath(), "gcsm-isolation-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var broken = Contact("broken");
                var healthy = Contact("healthy");
                var brokenKey = ContactContentFingerprint.Key("broken");
                var healthyKey = ContactContentFingerprint.Key("healthy");
                var corruptPath = Path.Combine(directory, brokenKey + ".json");
                File.WriteAllText(corruptPath, "null");
                var tracker = new ContactContentTracker(directory);
                tracker.SaveUpdate(healthyKey, "old", () => healthy.GoogleContact.ResourceName);
                var sync = Sync(broken, healthy);
                ContactsMatcher.SyncContacts(sync, (match, _) =>
                {
                    var key = match == broken ? brokenKey : healthyKey;
                    match.GoogleContactDirty = tracker.Check(key, match.GoogleContact.ResourceName, "new", false)
                        == ContactContentDecision.Update;
                });
                sync.SaveContacts(sync.Contacts, match => sync.SaveContact(match,
                    _ => tracker.SaveUpdate(healthyKey, "new", () => match.GoogleContact.ResourceName)));
                Assert.That(sync.ErrorCount, Is.EqualTo(1));
                Assert.That(sync.SyncedCount, Is.EqualTo(1));
                Assert.That(File.ReadAllText(corruptPath), Is.EqualTo("null"));
                Assert.That(tracker.Check(healthyKey, healthy.GoogleContact.ResourceName, "new", false),
                    Is.EqualTo(ContactContentDecision.Unchanged));
            }
            finally { Directory.Delete(directory, true); }
        }

        [Test]
        public void PendingBaselineSurvivesPreparationFailureAndRetries()
        {
            var directory = Path.Combine(Path.GetTempPath(), "gcsm-isolation-" + Guid.NewGuid().ToString("N"));
            try
            {
                var match = Contact("retry");
                var key = ContactContentFingerprint.Key("retry");
                var tracker = new ContactContentTracker(directory);
                tracker.SaveUpdate(key, "old", () => match.GoogleContact.ResourceName);
                tracker.SaveUpdate(key, "new", () => null);
                var path = Path.Combine(directory, key + ".json");
                var pending = File.ReadAllText(path);
                var sync = Sync(match);
                ContactsMatcher.SyncContacts(sync, (item, _) => throw new IOException("Photo unreadable"));
                sync.SaveContacts(sync.Contacts, _ => throw new Exception("Unexpected save"));
                Assert.That(File.ReadAllText(path), Is.EqualTo(pending));
                tracker = new ContactContentTracker(directory);
                ContactsMatcher.SyncContacts(sync, (item, _) =>
                {
                    item.GoogleContactDirty = tracker.Check(key, item.GoogleContact.ResourceName, "new", false)
                        == ContactContentDecision.Update;
                });
                sync.SaveContacts(sync.Contacts, item => sync.SaveContact(item,
                    _ => tracker.SaveUpdate(key, "new", () => item.GoogleContact.ResourceName)));
                Assert.That(match.PreparationFailed, Is.False);
                Assert.That(sync.SyncedCount, Is.EqualTo(1));
                Assert.That(tracker.Check(key, match.GoogleContact.ResourceName, "new", false),
                    Is.EqualTo(ContactContentDecision.Unchanged));
            }
            finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
        }

        [TestCase(false)]
        [TestCase(true)]
        public void CancellationStopsPreparationInsteadOfBeingReportedAsContactFailure(bool legacy)
        {
            var sync = Sync(Contact("cancelled"), Contact("later"));
            Exception cancellation = legacy ? (Exception)new ApplicationException("Cancelled") : new OperationCanceledException();
            var wrapped = new Exception("Context", cancellation);
            var calls = 0;
            var caught = Assert.Throws<Exception>(() => ContactsMatcher.SyncContacts(sync, (match, _) =>
            {
                calls++;
                throw wrapped;
            }));
            Assert.That(caught, Is.SameAs(wrapped));
            Assert.That(calls, Is.EqualTo(1));
            Assert.That(sync.ErrorCount, Is.Zero);
        }

        [Test]
        public void ReplacementsFromFailedPreparationAreDiscarded()
        {
            var broken = Contact("broken");
            var healthy = Contact("healthy");
            var sync = Sync(broken, healthy);
            ContactsMatcher.SyncContacts(sync, (match, current) =>
            {
                if (match != broken) return;
                current.Contacts.Add(new ContactMatch(null, match.GoogleContact));
                throw new IOException("Late read failure");
            });
            Assert.That(sync.Contacts, Is.EqualTo(new[] { broken, healthy }));
            Assert.That(healthy.PreparationFailed, Is.False);
            Assert.That(sync.ErrorCount, Is.EqualTo(1));
        }

        [Test]
        public void OutlookLookupFailureIsCaughtByTheRealPreparationLoop()
        {
            var match = Contact("missing cached Outlook ID");
            var sync = Sync(match);
            // GetOriginalItemFromOutlook rejects a missing EntryID before contacting Outlook.
            Assert.DoesNotThrow(() => ContactsMatcher.SyncContacts(sync));
            Assert.That(match.PreparationFailed, Is.True);
            Assert.That(sync.ErrorCount, Is.EqualTo(1));
            Assert.DoesNotThrow(() => sync.SaveContacts(sync.Contacts));
        }
    }
}
