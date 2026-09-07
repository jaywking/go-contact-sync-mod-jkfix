using Google.Apis.PeopleService.v1.Data;
using NUnit.Framework;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    [NonParallelizable] // The selected Google group is a static profile setting.
    public class GoogleGroupPreservationTests
    {
        private const string Starred = "contactGroups/starred";
        private Synchronizer sync;
        private string previousSelectedGroup;

        [SetUp]
        public void SetUp()
        {
            previousSelectedGroup = Synchronizer.SyncContactsGoogleGroup;
            Synchronizer.SyncContactsGoogleGroup = null;
            sync = new Synchronizer
            {
                SyncOption = SyncOption.OutlookToGoogleOnly,
                GoogleGroups = new Collection<ContactGroup>
                {
                    new ContactGroup { ResourceName = Starred, Name = "Favoriten" },
                    new ContactGroup { ResourceName = "contactGroups/work", Name = "Work" },
                    new ContactGroup { ResourceName = "contactGroups/old", Name = "Old" },
                    new ContactGroup { ResourceName = "contactGroups/custom", Name = "starred" }
                }
            };
        }

        [TearDown]
        public void TearDown()
        {
            Synchronizer.SyncContactsGoogleGroup = previousSelectedGroup;
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("Work")]
        public void OneWaySyncPreservesFavoriteAndReconcilesOtherCategories(string categories)
        {
            var contact = Contact(Starred, "contactGroups/work", "contactGroups/old", Synchronizer.myContactsGroup);

            sync.RemoveObsoleteGoogleGroups(contact, Utilities.GetOutlookGroups(categories));
            // A second sync must leave the same memberships, without duplication.
            sync.RemoveObsoleteGoogleGroups(contact, Utilities.GetOutlookGroups(categories));

            var expected = new List<string> { Starred, Synchronizer.myContactsGroup };
            if (categories == "Work")
                expected.Add("contactGroups/work");
            Assert.That(Resources(contact), Is.EquivalentTo(expected));
        }

        [Test]
        public void UserLabelNamedStarredIsStillRemoved()
        {
            var contact = Contact(Starred, "contactGroups/custom");
            sync.RemoveObsoleteGoogleGroups(contact, new string[0]);
            Assert.That(Resources(contact), Is.EqualTo(new[] { Starred }));
        }

        [Test]
        public void SelectedSyncGroupIsStillPreserved()
        {
            Synchronizer.SyncContactsGoogleGroup = "contactGroups/work";
            var contact = Contact(Starred, "contactGroups/work", "contactGroups/old");
            sync.RemoveObsoleteGoogleGroups(contact, new string[0]);
            Assert.That(Resources(contact), Is.EquivalentTo(new[] { Starred, "contactGroups/work" }));
        }

        [TestCase(SyncOption.MergeOutlookWins)]
        [TestCase(SyncOption.MergeGoogleWins)]
        [TestCase(SyncOption.MergePrompt)]
        public void TwoWaySyncStillRemovesFavoriteWhenCategoryIsRemoved(SyncOption option)
        {
            sync.SyncOption = option;
            var contact = Contact(Starred, "contactGroups/work");
            sync.RemoveObsoleteGoogleGroups(contact, new[] { "Work" });
            Assert.That(Resources(contact), Is.EqualTo(new[] { "contactGroups/work" }));
        }

        [Test]
        public void OneWaySyncDoesNotStarAnUnstarredContact()
        {
            var contact = Contact("contactGroups/work");
            sync.RemoveObsoleteGoogleGroups(contact, new[] { "Work" });
            Assert.That(Resources(contact), Is.EqualTo(new[] { "contactGroups/work" }));
        }

        [Test]
        public void ContactWithoutMembershipsIsUnchanged()
        {
            var contact = new Person();
            sync.RemoveObsoleteGoogleGroups(contact, new string[0]);
            Assert.That(contact.Memberships, Is.Null);
        }

        private static Person Contact(params string[] resources)
        {
            return new Person
            {
                Memberships = resources.Select(resource => new Membership
                {
                    ContactGroupMembership = new ContactGroupMembership { ContactGroupResourceName = resource }
                }).ToList()
            };
        }

        private static string[] Resources(Person contact)
        {
            return contact.Memberships.Select(m => m.ContactGroupMembership.ContactGroupResourceName).ToArray();
        }
    }
}
