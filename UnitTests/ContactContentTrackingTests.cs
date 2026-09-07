using Google.Apis.PeopleService.v1.Data;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class ContactContentTrackingTests
    {
        private string directory;
        private ContactContentTracker tracker;
        private readonly string key = ContactContentFingerprint.Key("profile", "account", "folder", "contact");
        private const string GoogleId = "people/test";

        [SetUp]
        public void SetUp()
        {
            directory = Path.Combine(Path.GetTempPath(), "gcsm-content-tests-" + Guid.NewGuid().ToString("N"));
            tracker = new ContactContentTracker(directory);
        }

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }

        private static string Company(string name) => Fingerprint(new Person
        {
            Organizations = new List<Organization> { new Organization { Name = name } }
        });

        private static string Fingerprint(Person person) => ContactContentFingerprint.Create(person, new[] { "Work" }, true, false, null);

        [Test]
        public void FirstObservationDoesNotWriteGoogleAndTracksTheNextImmediateEdit()
        {
            Assert.That(tracker.Check(key, GoogleId, Company("CBS Sports"), false), Is.EqualTo(ContactContentDecision.BaselineRecorded));
            tracker = new ContactContentTracker(directory); // Restart preserves the starting baseline.
            Assert.That(tracker.Check(key, GoogleId, Company("CBS Sports A"), false), Is.EqualTo(ContactContentDecision.Update));
        }

        [Test]
        public void ImmediateEditSyncsEvenWhenLegacyTimestampCheckSaysUnchanged()
        {
            var old = Company("CBS Sports");
            Assert.That(tracker.SaveUpdate(key, old, () => GoogleId), Is.True);
            var lastSync = new DateTime(2026, 9, 6, 21, 24, 0);
            var legacyChanged = Utilities.UpdatedSinceLastSync(lastSync.AddSeconds(18), lastSync);
            Assert.That(legacyChanged, Is.False);
            Assert.That(tracker.Check(key, GoogleId, Company("CBS Sports A"), legacyChanged), Is.EqualTo(ContactContentDecision.Update));
        }

        [Test]
        public void UnchangedContentsSkipEvenWhenTimestampsChange()
        {
            var fingerprint = Company("CBS Sports");
            tracker.SaveUpdate(key, fingerprint, () => GoogleId);
            Assert.That(tracker.Check(key, GoogleId, fingerprint, true), Is.EqualTo(ContactContentDecision.Unchanged));
            Assert.That(new ContactContentTracker(directory).Check(key, GoogleId, fingerprint, true), Is.EqualTo(ContactContentDecision.Unchanged));
        }

        [Test]
        public void ChangedLegacyContactStillUpdatesInsteadOfBeingBaselined()
        {
            Assert.That(tracker.Check(key, GoogleId, Company("new"), true), Is.EqualTo(ContactContentDecision.Update));
            Assert.That(Directory.Exists(directory), Is.False);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void FailedSaveStaysPendingAcrossRestart(bool existingBaseline)
        {
            if (existingBaseline)
                tracker.SaveUpdate(key, Company("old"), () => GoogleId);
            Assert.That(tracker.SaveUpdate(key, Company("new"), () => null), Is.False);
            tracker = new ContactContentTracker(directory);
            Assert.That(tracker.Check(key, GoogleId, Company("new"), false), Is.EqualTo(ContactContentDecision.Update));
            // Even reverting the edit must repair a potentially partial remote write.
            Assert.That(tracker.Check(key, GoogleId, Company("old"), false), Is.EqualTo(ContactContentDecision.Update));
            tracker.SaveUpdate(key, Company("new"), () => GoogleId);
            Assert.That(tracker.Check(key, GoogleId, Company("new"), false), Is.EqualTo(ContactContentDecision.Unchanged));
        }

        [Test]
        public void ExceptionAfterRemoteSaveDoesNotAdvanceBaseline()
        {
            tracker.SaveUpdate(key, Company("old"), () => GoogleId);
            Assert.Throws<IOException>(() => tracker.SaveUpdate(key, Company("new"), () =>
            {
                throw new IOException("Simulated photo or Outlook metadata save failure");
            }));
            Assert.That(new ContactContentTracker(directory).Check(key, GoogleId, Company("new"), false), Is.EqualTo(ContactContentDecision.Update));
        }

        [Test]
        public void EditDuringRequestIsNotMistakenForSavedContent()
        {
            var source = Company("first edit");
            tracker.SaveUpdate(key, source, () =>
            {
                Assert.That(new ContactContentTracker(directory).Check(key, GoogleId, source, false), Is.EqualTo(ContactContentDecision.Update));
                source = Company("second edit during request");
                return GoogleId;
            });
            Assert.That(tracker.Check(key, GoogleId, source, false), Is.EqualTo(ContactContentDecision.Update));
        }

        [Test]
        public void RelinkingGoogleContactForcesAnUpdate()
        {
            tracker.SaveUpdate(key, Company("same"), () => GoogleId);
            Assert.That(tracker.Check(key, "people/other", Company("same"), false), Is.EqualTo(ContactContentDecision.Update));
        }

        [Test]
        public void ProfilesAndAccountsHaveSeparateBaselines()
        {
            tracker.SaveUpdate(key, Company("old"), () => GoogleId);
            var other = ContactContentFingerprint.Key("profile", "other account", "folder", "contact");
            Assert.That(tracker.Check(other, GoogleId, Company("other"), false), Is.EqualTo(ContactContentDecision.BaselineRecorded));
            Assert.That(tracker.Check(key, GoogleId, Company("old"), false), Is.EqualTo(ContactContentDecision.Unchanged));
        }

        [Test]
        public void DiskContainsNoContactFieldValues()
        {
            tracker.SaveUpdate(key, Company("Private company name"), () => GoogleId);
            var contents = File.ReadAllText(Path.Combine(directory, key + ".json"));
            Assert.That(contents, Does.Not.Contain("Private company name"));
            Assert.That(contents, Does.Not.Contain(GoogleId));
        }

        [Test]
        public void CorruptBaselineFailsInsteadOfForgettingPendingEdits()
        {
            tracker.SaveUpdate(key, Company("old"), () => GoogleId);
            File.WriteAllText(Path.Combine(directory, key + ".json"), "null");
            Assert.Throws<InvalidDataException>(() => tracker.Check(key, GoogleId, Company("new"), false));
        }

        [Test]
        public void LocalPersistenceFailurePreventsRemoteWrite()
        {
            Directory.CreateDirectory(directory);
            var blocked = Path.Combine(directory, "file-not-directory");
            File.WriteAllText(blocked, "occupied");
            bool remoteCalled = false;
            Assert.Throws<IOException>(() => new ContactContentTracker(blocked).SaveUpdate(key, Company("new"), () =>
            {
                remoteCalled = true;
                return GoogleId;
            }));
            Assert.That(remoteCalled, Is.False);
        }

        [Test]
        public void MetadataAndGoogleFavoriteChangesDoNotChangeSourceFingerprint()
        {
            var person = new Person { Organizations = new List<Organization> { new Organization { Name = "CBS" } } };
            var before = Fingerprint(person);
            person.ETag = "new etag";
            person.ResourceName = "people/changed";
            person.ClientData = new List<ClientData> { new ClientData { Key = "sync", Value = "later" } };
            person.Organizations[0].Metadata = new FieldMetadata { Primary = true };
            person.Memberships = new List<Membership> { new Membership
            {
                ContactGroupMembership = new ContactGroupMembership { ContactGroupResourceName = "contactGroups/starred" }
            } };
            Assert.That(Fingerprint(person), Is.EqualTo(before));
        }

        [Test]
        public void CategoryOrderAndLegacyStarredSpellingAreStable()
        {
            var a = ContactContentFingerprint.Create(new Person(), new[] { "Work", "Starred in Android", "Work" }, true, false, null);
            var b = ContactContentFingerprint.Create(new Person(), new[] { "starred", "Work" }, true, false, null);
            Assert.That(a, Is.EqualTo(b));
            Assert.That(a, Is.Not.EqualTo(ContactContentFingerprint.Create(new Person(), new[] { "Work" }, true, false, null)));
        }

        [Test]
        public void PhotosAndMappingSettingsAreTracked()
        {
            var a = ContactContentFingerprint.Create(new Person(), null, true, true, "photo A");
            Assert.That(a, Is.Not.EqualTo(ContactContentFingerprint.Create(new Person(), null, true, true, "photo B")));
            Assert.That(a, Is.Not.EqualTo(ContactContentFingerprint.Create(new Person(), null, true, true, null)));
            Assert.That(a, Is.Not.EqualTo(ContactContentFingerprint.Create(new Person(), null, true, false, null)));
            Assert.That(a, Is.Not.EqualTo(ContactContentFingerprint.Create(new Person(), null, false, true, "photo A")));
        }

        [Test]
        public void NotesPreserveUnicodeButNormalizeLineEndings()
        {
            var a = Fingerprint(new Person { Biographies = new List<Biography> { new Biography { Value = "Équipe 😀\r\nLine 2" } } });
            var b = Fingerprint(new Person { Biographies = new List<Biography> { new Biography { Value = "Équipe 😀\nLine 2" } } });
            Assert.That(a, Is.EqualTo(b));
            Assert.That(a, Is.Not.EqualTo(Fingerprint(new Person { Biographies = new List<Biography> { new Biography { Value = "Équipe 😀\nNew text" } } })));
        }

        [Test]
        public void PhoneNameEmailAndAddressChangesAffectFingerprint()
        {
            var empty = Fingerprint(new Person());
            Assert.That(Fingerprint(new Person { PhoneNumbers = new List<PhoneNumber> { new PhoneNumber { Value = "5550100" } } }), Is.Not.EqualTo(empty));
            Assert.That(Fingerprint(new Person { Names = new List<Name> { new Name { GivenName = "Adam" } } }), Is.Not.EqualTo(empty));
            Assert.That(Fingerprint(new Person { EmailAddresses = new List<EmailAddress> { new EmailAddress { Value = "test@example.com" } } }), Is.Not.EqualTo(empty));
            Assert.That(Fingerprint(new Person { Addresses = new List<Address> { new Address { City = "New York" } } }), Is.Not.EqualTo(empty));
        }
    }
}
