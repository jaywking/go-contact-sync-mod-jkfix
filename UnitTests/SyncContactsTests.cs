using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Util.Store;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using Serilog;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class SyncContactsTests
    {
        private Synchronizer sync;
        private const int defaultWait = 5000;
        private const int defaultWaitTries = 10;

        //Constants for test contact
        internal const string TEST_CONTACT_NAME = "AN_OUTLOOK_TEST_CONTACT";
        internal const string TEST_SAVE_AS = "SaveAs";
        internal const string TEST_CONTACT_EMAIL = "email00@outlook.com";
        private const string TEST_GROUP = "A TEST GROUP";

        //private ContactGroup defaultContactGroup;
        private int initialGoogleGroupCount;
        private int initialGoogleContactCount;

        [OneTimeSetUp]
        public void Init()
        {
            //string timestamp = DateTime.Now.Ticks.ToString();            
            GoogleAPITests.LoadSettings(out var gmailUsername, out var syncProfile, out var syncContactsFolder, out var syncAppointmentsFolder);

            sync = new Synchronizer
            {
                SyncContacts = true,
                SyncAppointments = false,
            };
            Synchronizer.SyncProfile = syncProfile;
            Synchronizer.SyncContactsFolder = syncContactsFolder;

            sync.LoginToGoogle(gmailUsername);
            sync.LoginToOutlook();

            //Only load Google Contacts in My Contacts group (to avoid syncing accounts added automatically to "Weitere Kontakte"/"Further Contacts")
            sync.LoadGoogleGroups();
            //defaultContactGroup = sync.GetGoogleGroupByName(Synchronizer.myContactsGroup);

            initialGoogleGroupCount = CountGoogleGroups();
        }

        [SetUp]
        public void SetUp()
        {
            // delete previously failed test contacts
            DeleteTestContacts();
            DeleteTestContactGroups();
            initialGoogleGroupCount = CountGoogleGroups();
            sync.UseFileAs = true;
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            sync.LogoffOutlook();
            sync.LogoffGoogle();
        }

        [Test]
        public void TestLoadSingleContact()
        {
            Log.Information($"*** TestLoadSingleContact ***");

            LoadSettings(out var gmailUsername, out _);

            PeopleServiceService service;

            var scopes = new List<string>();
            //Contacts-Scope
            scopes.Add(PeopleServiceService.Scope.Contacts);
            //Calendar-Scope
            scopes.Add(CalendarService.Scope.Calendar);

            UserCredential credential;
            var jsonSecrets = Properties.Resources.client_secrets;

            using (var stream = new MemoryStream(jsonSecrets))
            {
                var Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\GoContactSyncMOD\\";
                var AuthFolder = Folder + "\\Auth\\";

                var fDS = new FileDataStore(AuthFolder, true);

                var clientSecrets = GoogleClientSecrets.FromStream(stream);

                credential = GCSMOAuth2WebAuthorizationBroker.AuthorizeAsync(
                                clientSecrets.Secrets,
                                scopes.ToArray(),
                                gmailUsername,
                                CancellationToken.None,
                                fDS).
                                Result;
                var initializer = new Google.Apis.Services.BaseClientService.Initializer
                {
                    HttpClientInitializer = credential
                };

                service = GoogleServices.CreatePeopleService(initializer);
            }

            #region Delete previously created test contact.
            var peopleResource = new PeopleResource(service);
            var peopleRequest = peopleResource.Connections.List("people/me");
            peopleRequest.PersonFields = Synchronizer.GET_PERSON_FIELDS;
            peopleRequest.PageToken = null;

            do
            {
                var response = peopleRequest.Execute();
                if (response.Connections != null)
                {
                    foreach (var entry in response.Connections)
                    {
                        var pm = ContactPropertiesUtils.GetGooglePrimaryEmail(entry);
                        if (pm != null && pm.Value == SyncContactsTests.TEST_CONTACT_EMAIL)
                        {
                            peopleResource.DeleteContact(entry.ResourceName).Execute();
                            Log.Information("Deleted Google contact");
                        }

                        peopleRequest.PageToken = response.NextPageToken;
                    }
                }
            } while (!string.IsNullOrEmpty(peopleRequest.PageToken));
            #endregion

            var newEntry = new GoogleContactBuilder().Build("Who is this guy?");

            var createdEntry = peopleResource.CreateContact(newEntry).Execute();

            Log.Information("Created Google contact");

            ClassicAssert.IsNotNull(createdEntry.ResourceName);

            var ret = sync.LoadGoogleContacts(ContactPropertiesUtils.GetGoogleId(createdEntry));
            ClassicAssert.IsNotNull(ret);

            peopleResource.DeleteContact(createdEntry.ResourceName).Execute();

            Log.Information($"*** TestLoadSingleContact ***");
        }

        [Test]
        public void TestSync_Structured()
        {
            Log.Information($"*** TestSync_Structured ***");

            // create new contact to sync
            var oc1 = new OutlookContactBuilder().Build();

            oc1.FileAs = TEST_CONTACT_NAME;
            oc1.Email1Address = TEST_CONTACT_EMAIL;
            oc1.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc1.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");

            oc1.HomeAddressStreet = "Street";
            oc1.HomeAddressCity = "City";
            oc1.HomeAddressPostalCode = "1234";
            oc1.HomeAddressCountry = "Country";
            oc1.HomeAddressPostOfficeBox = "PO1";
            oc1.HomeAddressState = "State1";

            oc1.BusinessAddressStreet = "Street2";
            oc1.BusinessAddressCity = "City2";
            oc1.BusinessAddressPostalCode = "5678";
            oc1.BusinessAddressCountry = "Country2";
            oc1.BusinessAddressPostOfficeBox = "PO2";
            oc1.BusinessAddressState = "State2";

            oc1.OtherAddressStreet = "Street3";
            oc1.OtherAddressCity = "City3";
            oc1.OtherAddressPostalCode = "8012";
            oc1.OtherAddressCountry = "Country3";
            oc1.OtherAddressPostOfficeBox = "PO3";
            oc1.OtherAddressState = "State3";

            #region phones
            //First delete the destination phone numbers
            oc1.PrimaryTelephoneNumber = "123";
            oc1.HomeTelephoneNumber = "456";
            oc1.Home2TelephoneNumber = "4567";
            oc1.BusinessTelephoneNumber = "45678";
            oc1.Business2TelephoneNumber = "456789";
            oc1.MobileTelephoneNumber = "123";
            oc1.BusinessFaxNumber = "1234";
            oc1.HomeFaxNumber = "12345";
            oc1.PagerNumber = "123456";
            oc1.OtherTelephoneNumber = "12345678";
            oc1.CarTelephoneNumber = "123456789";
            oc1.AssistantTelephoneNumber = "987";
            #endregion phones

            #region Name
            oc1.Title = "Title";
            oc1.FirstName = "Firstname";
            oc1.MiddleName = "Middlename";
            oc1.LastName = "Lastname";
            oc1.Suffix = "Suffix";
            #endregion Name

            oc1.Birthday = new DateTime(1999, 1, 1);

            oc1.NickName = "Nickname";
            oc1.OfficeLocation = "Location";
            oc1.Initials = "IN";
            oc1.Language = "German";

            oc1.CompanyName = "CompanyName";
            oc1.JobTitle = "Position";
            oc1.Department = "Department";

            oc1.IMAddress = "IMs";
            oc1.Anniversary = new DateTime(2000, 1, 1);
            oc1.Children = "Children";
            oc1.Spouse = "Spouse";
            oc1.AssistantName = "Assi";
            oc1.ManagerName = "Chef";
            oc1.WebPage = "http://www.test.de";
            oc1.Body = "<sn>Content & other stuff</sn>\r\n<![CDATA[  \r\n...\r\n&stuff in CDATA < >\r\n  ]]>";
            oc1.Save();

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc1, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc1, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveGoogleContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            //ContactsMatcher.SyncContact(match, sync);

            var oc2 = new OutlookContactBuilder().Build();
            ContactSync.UpdateContact(match.GoogleContact, oc2, sync.UseFileAs);

            // match oc2 with oc1
            ClassicAssert.AreEqual(oc1.FileAs, oc2.FileAs);
            ClassicAssert.AreEqual(oc1.Email1Address, oc2.Email1Address);
            ClassicAssert.AreEqual(oc1.Email2Address, oc2.Email2Address);
            ClassicAssert.AreEqual(oc1.Email3Address, oc2.Email3Address);
            ClassicAssert.AreEqual(oc1.PrimaryTelephoneNumber, oc2.PrimaryTelephoneNumber);
            ClassicAssert.AreEqual(oc1.HomeTelephoneNumber, oc2.HomeTelephoneNumber);
            ClassicAssert.AreEqual(oc1.Home2TelephoneNumber, oc2.Home2TelephoneNumber);
            ClassicAssert.AreEqual(oc1.BusinessTelephoneNumber, oc2.BusinessTelephoneNumber);
            ClassicAssert.AreEqual(oc1.Business2TelephoneNumber, oc2.Business2TelephoneNumber);
            ClassicAssert.AreEqual(oc1.MobileTelephoneNumber, oc2.MobileTelephoneNumber);
            ClassicAssert.AreEqual(oc1.BusinessFaxNumber, oc2.BusinessFaxNumber);
            ClassicAssert.AreEqual(oc1.HomeFaxNumber, oc2.HomeFaxNumber);
            ClassicAssert.AreEqual(oc1.PagerNumber, oc2.PagerNumber);
            ClassicAssert.AreEqual(oc1.OtherTelephoneNumber, oc2.OtherTelephoneNumber);

            ClassicAssert.AreEqual(oc1.HomeAddressStreet, oc2.HomeAddressStreet);
            ClassicAssert.AreEqual(oc1.HomeAddressCity, oc2.HomeAddressCity);
            ClassicAssert.AreEqual(oc1.HomeAddressCountry, oc2.HomeAddressCountry);
            ClassicAssert.AreEqual(oc1.HomeAddressPostalCode, oc2.HomeAddressPostalCode);
            ClassicAssert.AreEqual(oc1.HomeAddressPostOfficeBox, oc2.HomeAddressPostOfficeBox);
            ClassicAssert.AreEqual(oc1.HomeAddressState, oc2.HomeAddressState);

            ClassicAssert.AreEqual(oc1.BusinessAddressStreet, oc2.BusinessAddressStreet);
            ClassicAssert.AreEqual(oc1.BusinessAddressCity, oc2.BusinessAddressCity);
            ClassicAssert.AreEqual(oc1.BusinessAddressCountry, oc2.BusinessAddressCountry);
            ClassicAssert.AreEqual(oc1.BusinessAddressPostalCode, oc2.BusinessAddressPostalCode);
            ClassicAssert.AreEqual(oc1.BusinessAddressPostOfficeBox, oc2.BusinessAddressPostOfficeBox);
            ClassicAssert.AreEqual(oc1.BusinessAddressState, oc2.BusinessAddressState);

            ClassicAssert.AreEqual(oc1.OtherAddressStreet, oc2.OtherAddressStreet);
            ClassicAssert.AreEqual(oc1.OtherAddressCity, oc2.OtherAddressCity);
            ClassicAssert.AreEqual(oc1.OtherAddressCountry, oc2.OtherAddressCountry);
            ClassicAssert.AreEqual(oc1.OtherAddressPostalCode, oc2.OtherAddressPostalCode);
            ClassicAssert.AreEqual(oc1.OtherAddressPostOfficeBox, oc2.OtherAddressPostOfficeBox);
            ClassicAssert.AreEqual(oc1.OtherAddressState, oc2.OtherAddressState);

            ClassicAssert.AreEqual(oc1.FullName, oc2.FullName);
            ClassicAssert.AreEqual(oc1.MiddleName, oc2.MiddleName);
            ClassicAssert.AreEqual(oc1.LastName, oc2.LastName);
            ClassicAssert.AreEqual(oc1.FirstName, oc2.FirstName);
            ClassicAssert.AreEqual(oc1.Title, oc2.Title);
            ClassicAssert.AreEqual(oc1.Suffix, oc2.Suffix);

            //Expected: 1999 - 01 - 01 00:00:00
            //But was:  4501 - 01 - 01 00:00:00
            //TODO, sync is not working due to issue at Google side, see:
            //https://issuetracker.google.com/issues/200184636
            //ClassicAssert.AreEqual(oc1.Birthday, oc2.Birthday);

            ClassicAssert.AreEqual(oc1.NickName, oc2.NickName);
            ClassicAssert.AreEqual(oc1.OfficeLocation, oc2.OfficeLocation);
            ClassicAssert.AreEqual(oc1.Initials, oc2.Initials);
            ClassicAssert.AreEqual(oc1.Language, oc2.Language);

            ClassicAssert.AreEqual(oc1.IMAddress, oc2.IMAddress);
            ClassicAssert.AreEqual(oc1.Anniversary, oc2.Anniversary);
            ClassicAssert.AreEqual(oc1.Children, oc2.Children);
            ClassicAssert.AreEqual(oc1.Spouse, oc2.Spouse);
            ClassicAssert.AreEqual(oc1.ManagerName, oc2.ManagerName);
            ClassicAssert.AreEqual(oc1.AssistantName, oc2.AssistantName);

            ClassicAssert.AreEqual(oc1.WebPage, oc2.WebPage);
            ClassicAssert.AreEqual(oc1.Body, oc2.Body);

            ClassicAssert.AreEqual(oc1.CompanyName, oc2.CompanyName);
            ClassicAssert.AreEqual(oc1.JobTitle, oc2.JobTitle);
            ClassicAssert.AreEqual(oc1.Department, oc2.Department);

            DeleteAppointmentsForTestContacts();

            oc1.Delete();
            oc2.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSync_Structured ***");
        }

        [Test]
        public void TestSync_Unstructured()
        {
            Log.Information($"*** TestSync_Unstructured ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var oc1 = new OutlookContactBuilder().Build();
            oc1.FileAs = TEST_CONTACT_NAME;
            oc1.HomeAddress = "10 Parades";
            oc1.BusinessAddress = "11 Parades";
            oc1.OtherAddress = "12 Parades";
            oc1.IMAddress = "  "; //Test empty IMAddress
            oc1.Email2Address = "  "; //Test empty Email Address
            oc1.FullName = TEST_CONTACT_NAME;
            oc1.Save();

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc1, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc1, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveGoogleContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            //ContactsMatcher.SyncContact(match, sync);
            ClassicAssert.IsNotNull(match.GoogleContact);

            var oc2 = new OutlookContactBuilder().Build();
            ContactSync.UpdateContact(match.GoogleContact, oc2, sync.UseFileAs);

            // match oc2 with oc1
            ClassicAssert.AreEqual(oc1.FileAs, oc2.FileAs);

            ClassicAssert.AreEqual(oc1.HomeAddress, oc2.HomeAddress);
            ClassicAssert.AreEqual(oc1.BusinessAddress, oc2.BusinessAddress);
            ClassicAssert.AreEqual(oc1.OtherAddress, oc2.OtherAddress);

            ClassicAssert.AreEqual(oc1.FullName, oc2.FullName);

            oc1.Delete();
            oc2.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSync_Unstructured ***");
        }

        [Test]
        public void TestSync_CompanyOnly()
        {
            Log.Information($"*** TestSync_CompanyOnly ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var oc1 = new OutlookContactBuilder().Build();
            oc1.CompanyName = TEST_CONTACT_NAME;
            oc1.BusinessAddress = "11 Parades";
            oc1.Save();

            ClassicAssert.IsNull(oc1.FullName);
            ClassicAssert.IsNull(oc1.Email1Address);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc1, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc1, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveGoogleContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            //ContactsMatcher.SyncContact(match, sync);

            var oc2 = new OutlookContactBuilder().Build();
            ContactSync.UpdateContact(match.GoogleContact, oc2, sync.UseFileAs);

            // match oc2 with oc1
            ClassicAssert.AreEqual(oc1.FileAs, oc2.FileAs);
            ClassicAssert.AreEqual(oc1.CompanyName, oc2.CompanyName);
            ClassicAssert.AreEqual(oc1.BusinessAddress, oc2.BusinessAddress);
            ClassicAssert.IsNull(oc2.FullName);
            ClassicAssert.IsNull(oc1.Email1Address);

            oc1.Delete();
            oc2.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSync_CompanyOnly ***");
        }

        [Test]
        public void TestSync_EmailOnly()
        {
            Log.Information($"*** TestSync_EmailOnly ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var oc1 = new OutlookContactBuilder().Build();
            oc1.FileAs = TEST_CONTACT_EMAIL;
            oc1.Email1Address = TEST_CONTACT_EMAIL;
            oc1.Save();

            ClassicAssert.IsNull(oc1.FullName);
            ClassicAssert.IsNull(oc1.CompanyName);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc1, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc1, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveGoogleContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_EMAIL, TEST_CONTACT_EMAIL);
            //ContactsMatcher.SyncContact(match, sync);

            var oc2 = new OutlookContactBuilder().Build();
            ContactSync.UpdateContact(match.GoogleContact, oc2, sync.UseFileAs);

            // match oc2 with oc1
            ClassicAssert.AreEqual(oc1.FileAs, oc2.FileAs);
            ClassicAssert.AreEqual(oc1.Email1Address, oc2.Email1Address);
            ClassicAssert.IsNull(oc2.FullName);
            ClassicAssert.IsNull(oc1.CompanyName);

            oc1.Delete();
            oc2.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSync_EmailOnly ***");
        }

        [Test]
        public void TestSync_UseFileAs()
        {
            Log.Information($"*** TestSync_UseFileAs ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.UseFileAs = true;

            // create new contact to sync
            var oc1 = new OutlookContactBuilder().Build();
            oc1.FullName = TEST_CONTACT_NAME;
            oc1.FileAs = TEST_SAVE_AS;
            oc1.Save();

            ClassicAssert.AreNotEqual(oc1.FullName, oc1.FileAs);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc1, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc1, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveGoogleContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_SAVE_AS, string.Empty);
            //ContactsMatcher.SyncContact(match, sync);

            var oc2 = new OutlookContactBuilder().Build();
            ClassicAssert.IsNotNull(match.GoogleContact);
            ContactSync.UpdateContact(match.GoogleContact, oc2, sync.UseFileAs);

            // match oc2 with oc1
            var googleName = ContactPropertiesUtils.GetGooglePrimaryName(match.GoogleContact);
            ClassicAssert.IsNotNull(googleName);
            var googleFileAs = ContactPropertiesUtils.GetGoogleFileAs(match.GoogleContact);
            ClassicAssert.IsNotNull(googleFileAs);

            ClassicAssert.AreEqual(oc2.FileAs, googleFileAs.Value);
            ClassicAssert.AreEqual(oc2.FileAs, googleName.UnstructuredName);
            ClassicAssert.AreEqual(oc1.FileAs, oc2.FileAs);

            oc2.FileAs = TEST_CONTACT_NAME;
            ClassicAssert.AreNotEqual(oc1.FileAs, oc2.FileAs);
            ClassicAssert.AreNotEqual(oc2.FileAs, googleFileAs.Value);
            ContactSync.UpdateContact(match.GoogleContact, oc2, sync.UseFileAs);
            ClassicAssert.AreEqual(googleName.FamilyName, oc2.FileAs);

            oc1.Delete();
            oc2.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSync_UseFileAs ***");
        }

        [Test]
        public void TestSync_UseFullName()
        {
            Log.Information($"*** TestSync_UseFullName ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.UseFileAs = false;

            // create new contact to sync
            var oc1 = new OutlookContactBuilder().Build();
            oc1.FullName = TEST_CONTACT_NAME;
            oc1.FileAs = TEST_SAVE_AS;
            oc1.Save();

            ClassicAssert.AreNotEqual(oc1.FullName, oc1.FileAs);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc1, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc1, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveGoogleContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNotNull(match);

            var oc2 = new OutlookContactBuilder().Build();
            ContactSync.UpdateContact(match.GoogleContact, oc2, sync.UseFileAs);

            // match oc2 with oc1
            var googleName = ContactPropertiesUtils.GetGooglePrimaryName(match.GoogleContact);
            ClassicAssert.IsNotNull(googleName);
            Assert.That(googleName.UnstructuredName, Is.Not.Empty);
            var googleFileAs = ContactPropertiesUtils.GetGoogleFileAs(match.GoogleContact);
            ClassicAssert.IsNull(googleFileAs);

            ClassicAssert.AreEqual(googleName.UnstructuredName, OutlookContactInfo.GetTitleFirstLastAndSuffix(oc2).Trim().Replace("  ", " "));
            //ClassicAssert.AreNotEqual(oc1.FileAs, googleFileAs.Value);
            ClassicAssert.AreNotEqual(oc1.FileAs, googleName.UnstructuredName);
            ClassicAssert.AreNotEqual(oc1.FileAs, oc2.FileAs);

            oc2.FileAs = TEST_SAVE_AS;
            ClassicAssert.AreEqual(oc1.FileAs, oc2.FileAs);
            ContactSync.UpdateContact(match.GoogleContact, oc2, sync.UseFileAs);
            ClassicAssert.AreEqual(oc1.FileAs, oc2.FileAs);

            oc1.Delete();
            oc2.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSync_UseFullName ***");
        }

        [Test]
        public void TestExtendedProps()
        {
            Log.Information($"*** TestExtendedProps ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.UseFileAs = true;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var m = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            m.GoogleContactDirty = true;

            sync.SaveGoogleContact(m);

            var googleFileAs = ContactPropertiesUtils.GetGoogleFileAs(m.GoogleContact);
            ClassicAssert.IsNotNull(googleFileAs);

            ClassicAssert.AreEqual(TEST_CONTACT_NAME, googleFileAs.Value);

            // read contact from google
            Assert.That(EnsureGoogleContactSaved(m.GoogleContact));
            sync.MatchContacts();
            ContactsMatcher.SyncContacts(sync);

            m = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);

            ClassicAssert.IsNotNull(m);
            ClassicAssert.IsNotNull(m.GoogleContact);

            // get extended prop
            var oid = ContactPropertiesUtils.GetOutlookId(oc);
            var gid = ContactPropertiesUtils.GetGoogleOutlookContactId(m.GoogleContact);
            ClassicAssert.AreEqual(oid, gid);

            oc.Delete();

            DeleteTestContact(m.GoogleContact);

            Log.Information($"*** TestExtendedProps ***");
        }

        [Test]
        public void TestSyncDeletedOulook()
        {
            Log.Information($"*** TestSyncDeletedOulook ***");

            sync.LoadContacts();

            //ToDo: Check for each SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contacts
            sync.SaveContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            // delete outlook contact
            oc.Delete();

            // sync
            sync.MatchContacts();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNotNull(match);
            ClassicAssert.IsNotNull(match.GoogleContact);
            ClassicAssert.IsNull(match.OutlookContact);

            // delete
            sync.SaveContact(match);
            Assert.That(EnsureGoogleContactDeleted(match.GoogleContact));

            // sync
            sync.MatchContacts();
            ContactsMatcher.SyncContacts(sync);

            // check if google contact still exists
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);

            ClassicAssert.IsNull(match);

            Log.Information($"*** TestSyncDeletedOulook ***");
        }

        [Test]
        public void TestSyncDeletedGoogle()
        {
            Log.Information($"*** TestSyncDeletedGoogle ***");

            //ToDo: Check for each SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contacts
            sync.SaveContact(match);

            // delete google contact
            sync.GooglePeopleResource.DeleteContact(match.GoogleContact.ResourceName).Execute();

            // sync
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ContactsMatcher.SyncContact(match, sync);

            // delete
            sync.SaveContact(match);

            // sync
            sync.MatchContacts();
            ContactsMatcher.SyncContacts(sync);
            _ = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);

            Log.Information($"*** TestSyncDeletedGoogle ***");
        }

        [Test]
        public void TestGooglePhoto()
        {
            Log.Information($"*** TestGooglePhoto ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;
            Synchronizer.SyncPhotos = true;

            Assert.That(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveGoogleContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));
            ClassicAssert.IsNotNull(match.GoogleContact.ResourceName);

            //load the same contact from google.
            sync.MatchContacts();

            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNotNull(match.GoogleContact);
            ClassicAssert.IsNotNull(match.GoogleContact.ResourceName);
            ContactsMatcher.SyncContact(match, sync);

            var pic = Utilities.CropImageGoogleFormat(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));
            ClassicAssert.IsNotNull(match.GoogleContact.ResourceName);
            var saved = sync.SaveGooglePhoto(match.GoogleContact, new Bitmap(pic));
            Assert.That(saved);
            Assert.That(EnsureGoogleContactHasPhoto(match.GoogleContact));

            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ContactsMatcher.SyncContact(match, sync);

            var image = Utilities.GetGoogleContactPhoto(sync, match.GoogleContact);
            ClassicAssert.IsNotNull(image);

            oc.Delete();
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestGooglePhoto ***");
        }

        [Test]
        public void TestOutlookPhoto()
        {
            Log.Information($"*** TestOutlookPhoto ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;
            Synchronizer.SyncPhotos = true;

            Assert.That(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();

            var saved = oc.SetOutlookPhoto(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg");
            Assert.That(saved);

            oc.Save();

            var image = oc.GetOutlookPhoto();
            ClassicAssert.IsNotNull(image);

            oc.Delete();

            Log.Information($"*** TestOutlookPhoto ***");
        }

        [Test]
        public void TestSyncPhoto()
        {
            Log.Information($"*** TestSyncPhoto ***");

            sync.LoadContacts();

            sync.SyncOption = SyncOption.MergeOutlookWins;
            Synchronizer.SyncPhotos = true;

            Assert.That(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.SetOutlookPhoto(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg");
            oc.Save();

            // outlook contact should now have a photo
            ClassicAssert.IsNotNull(oc.GetOutlookPhoto());

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ContactsMatcher.SyncContact(match, sync);

            // google contact should now have a photo
            ClassicAssert.IsNotNull(Utilities.GetGoogleContactPhoto(sync, match.GoogleContact));

            // delete outlook contact
            oc.Delete();
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }

            //DeleteTestPerson(match.GoogleContact);

            // recreate outlook contact
            oc = new OutlookContactBuilder().Build();

            // outlook contact should now have no photo
            ClassicAssert.IsNull(oc.GetOutlookPhoto());

            sync.UpdateContact(match.GoogleContact, oc, match.GoogleContactDirty, match.matchedById);
            match = new ContactMatch(new OutlookContactInfo(oc, sync), match.GoogleContact);
            //match.OutlookContact.Save();            

            //save contact to outlook
            sync.SaveContact(match);

            // outlook contact should now have a photo
            ClassicAssert.IsNotNull(oc.GetOutlookPhoto());

            oc.Delete();
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSyncPhoto ***");
        }

        [Test]
        public void TestSyncLargePerson()
        {
            Log.Information($"*** TestSyncLargePerson ***");

            sync.LoadContacts();
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Body = new string('*', 150000);
            oc.Save();

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveContact(match);
            ClassicAssert.IsNull(match.GoogleContact);

            oc.Delete();
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }

            Log.Information($"*** TestSyncLargePerson ***");
        }

        [Test]
        public void TestSyncContactModifiedToBeLarge()
        {
            Log.Information($"*** TestSyncContactModifiedToBeLarge ***");

            //sync.LoadContacts();
            //Arrange
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();
            var oid = ContactPropertiesUtils.GetOutlookId(oc);
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }
            sync.MatchContacts();
            ContactsMatcher.SyncContacts(sync);
            sync.SaveContacts(sync.Contacts);

            DelayBetweenSync();

            oc = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            oc.Body = new string('*', 63000);
            oc.Save();
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }

            DelayBetweenSync();

            //Act
            var before = sync.GoogleContacts.Count;
            sync.MatchContacts();
            ContactsMatcher.SyncContacts(sync);
            sync.SaveContacts(sync.Contacts);
            //Assert.GreaterOrEqual(sync.Contacts.Count, 1);
            ClassicAssert.AreEqual(before+1, sync.Contacts.Count);
            ClassicAssert.AreEqual(1, sync.Contacts.Count - sync.GoogleContacts.Count);

            //Assert
            ClassicAssert.IsNotNull(oid);
            oc = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            ClassicAssert.IsNotNull(oc);
            ClassicAssert.AreEqual(new string('*', 63000), oc.Body);
            var gid = ContactPropertiesUtils.GetOutlookGoogleId(oc);
            ClassicAssert.IsNotNull(gid);
            var gc = RetrievePerson(gid);
            ClassicAssert.IsNotNull(gc);
            var bio = ContactPropertiesUtils.GetGoogleBiography(gc);
            ClassicAssert.IsNotNull(bio);
            ClassicAssert.AreEqual(new string('*', 63000), bio.Value);

            DeleteTestContact(oc);
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }
            DeleteTestContact(gc);

            Log.Information($"*** TestSyncContactModifiedToBeLarge ***");
        }

        [Test]
        public void TestSyncContactModifiedToBeLargeAndFail()
        {
            Log.Information($"*** TestSyncContactModifiedToBeLargeAndFail ***");

            //sync.LoadContacts();
            //Arrange
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();
            var oid = ContactPropertiesUtils.GetOutlookId(oc);
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }
            sync.MatchContacts();
            ContactsMatcher.SyncContacts(sync);
            sync.SaveContacts(sync.Contacts);

            DelayBetweenSync();

            oc = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            oc.Body = new string('*', 150000);
            oc.Save();
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }

            DelayBetweenSync();

            //Act
            var before = sync.GoogleContacts.Count;
            sync.MatchContacts();
            ContactsMatcher.SyncContacts(sync);
            sync.SaveContacts(sync.Contacts);
            //ClassicAssert.AreEqual(1, sync.Contacts.Count);
            //ClassicAssert.AreEqual(0, sync.GoogleContacts.Count);
            ClassicAssert.AreEqual(before+1, sync.Contacts.Count);
            ClassicAssert.AreEqual(1, sync.Contacts.Count - sync.GoogleContacts.Count);

            //Assert
            ClassicAssert.IsNotNull(oid);
            oc = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            ClassicAssert.IsNotNull(oc);
            ClassicAssert.AreEqual(new string('*', 150000), oc.Body);
            var gid = ContactPropertiesUtils.GetOutlookGoogleId(oc);
            ClassicAssert.IsNotNull(gid);
            var gc = RetrievePerson(gid);
            ClassicAssert.IsNotNull(gc);
            var bio = ContactPropertiesUtils.GetGoogleBiography(gc);
            ClassicAssert.IsNull(bio);

            DeleteTestContact(oc);
            if (oc != null)
            {
                Marshal.ReleaseComObject(oc);
            }
            DeleteTestContact(gc);

            Log.Information($"*** TestSyncContactModifiedToBeLargeAndFail ***");
        }

        [Test]
        public void TestSyncGroups()
        {
            Log.Information($"*** TestSyncGroups ***");

            Assert.That(EnsureGoogleGroupsCount(initialGoogleGroupCount));

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Categories = TEST_GROUP;
            oc.Save();

            //Outlook contact should now have a group
            ClassicAssert.AreEqual(TEST_GROUP, oc.Categories);

            //Sync GoogleGroups first
            sync.MatchContacts();
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.Contacts.Count);

            ClassicAssert.AreEqual(initialGoogleGroupCount, CountGoogleGroups());
            ContactsMatcher.SyncGroups(sync);
            Assert.That(EnsureGoogleGroupExist(TEST_GROUP));
            ClassicAssert.AreEqual(initialGoogleGroupCount + 1, CountGoogleGroups());

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            ClassicAssert.AreEqual(1, gc.Memberships.Count);

            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //sync and save contact to google.
            ContactsMatcher.SyncContact(match, sync);
            ClassicAssert.AreEqual(1, match.GoogleContact.Memberships.Count);

            sync.SaveContact(match);

            Assert.That(EnsureGoogleContactHasGroups(match.GoogleContact, 1));

            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ContactsMatcher.SyncContact(match, sync);

            // google contact should now have the same group
            Assert.That(EnsureGoogleContactHasGroups(match.GoogleContact, 1));
            var googleContactGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            ClassicAssert.AreEqual(1, googleContactGroups.Count);
            ClassicAssert.Contains(sync.GetGoogleGroupByName(TEST_GROUP), googleContactGroups);
            //ClassicAssert.Contains(sync.GetGoogleGroupByName(Synchronizer.myContactsGroup), googleContactGroups);

            // delete outlook contact
            oc.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSyncGroups ***");
        }

        [Test]
        public void TestSyncDeletedGoogleContactGroup()
        {
            Log.Information($"*** TestSyncDeletedGoogleContactGroup ***");

            //ToDo: Check for each SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Categories = TEST_GROUP;
            oc.Save();

            //Outlook contact should now have a group
            ClassicAssert.AreEqual(TEST_GROUP, oc.Categories);

            //Sync ContactGroups first
            sync.MatchContacts();
            ContactsMatcher.SyncGroups(sync);
            Assert.That(EnsureGoogleGroupExist(TEST_GROUP));

            //Create now Google Person and assing new ContactGroup
            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.            
            sync.SaveContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));
            ClassicAssert.AreEqual(2, match.GoogleContact.Memberships.Count);

            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNotNull(match.GoogleContact);
            ClassicAssert.IsNotNull(match.OutlookContact);
            ClassicAssert.AreEqual(2, match.GoogleContact.Memberships.Count);
            ContactsMatcher.SyncContact(match, sync);

            // google contact should now have the same group
            var googleContactGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            ClassicAssert.AreEqual(1, googleContactGroups.Count);

            var group = sync.GetGoogleGroupByName(TEST_GROUP);
            ClassicAssert.Contains(group, googleContactGroups);
            //ClassicAssert.Contains(sync.GetGoogleGroupByName(Synchronizer.myContactsGroup), googleContactGroups);

            // delete group from google contact
            Utilities.RemoveGoogleGroup(match.GoogleContact, group);

            googleContactGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            ClassicAssert.AreEqual(0, googleContactGroups.Count);
            //ClassicAssert.Contains(sync.GetGoogleGroupByName(Synchronizer.myContactsGroup), googleContactGroups);

            //save contact to google.
            sync.SaveGoogleContact(match.GoogleContact);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));
            Assert.That(EnsureGoogleContactHasGroups(match.GoogleContact, 0));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;

            //Sync ContactGroups first
            sync.MatchContacts();
            ContactsMatcher.SyncGroups(sync);

            //sync and save contact to outlook.
            var etag = match.GoogleContact.ETag;
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            sync.UpdateContact(match.GoogleContact, oc, match.GoogleContactDirty, match.matchedById);
            sync.SaveContact(match);
            Assert.That(EnsureGoogleContactUpdated(match.GoogleContact, etag));
            ClassicAssert.AreEqual(1, match.GoogleContact.Memberships.Count);

            // google and outlook should now have no category except for the System ContactGroup: My Contacts
            googleContactGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            ClassicAssert.AreEqual(0, googleContactGroups.Count);
            ClassicAssert.AreEqual(null, oc.Categories);
            //ClassicAssert.Contains(sync.GetGoogleGroupByName(Synchronizer.myContactsGroup), googleContactGroups);

            // delete test group
            if (group != null)
            {
                var groupsResource = new ContactGroupsResource(sync.GooglePeopleService);
                groupsResource.Delete(group.ResourceName);
            }

            oc.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSyncDeletedGoogleContactGroup ***");
        }

        [Test]
        public void TestSyncDeletedOutlookContactGroup()
        {
            Log.Information($"*** TestSyncDeletedOutlookContactGroup ***");

            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Categories = TEST_GROUP;
            oc.Save();

            //Outlook contact should now have a group
            ClassicAssert.AreEqual(TEST_GROUP, oc.Categories);

            //Now sync ContactGroups
            sync.MatchContacts();
            ContactsMatcher.SyncGroups(sync);

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ContactsMatcher.SyncContact(match, sync);

            // google contact should now have the same group
            var googleContactGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            var group = sync.GetGoogleGroupByName(TEST_GROUP);
            ClassicAssert.AreEqual(1, googleContactGroups.Count);
            //ClassicAssert.Contains(sync.GetGoogleGroupByName(Synchronizer.myContactsGroup), googleContactGroups);
            ClassicAssert.Contains(group, googleContactGroups);

            // delete group from outlook
            Utilities.RemoveOutlookGroup(oc, TEST_GROUP);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.MatchContacts();
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            sync.UpdateContact(oc, match.GoogleContact, match);

            // google and outlook should now have no category
            googleContactGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            ClassicAssert.AreEqual(null, oc.Categories);
            ClassicAssert.AreEqual(0, googleContactGroups.Count);
            //ClassicAssert.Contains(sync.GetGoogleGroupByName(Synchronizer.myContactsGroup), googleContactGroups);

            // delete test group
            if (group != null)
            {
                var groupsResource = new ContactGroupsResource(sync.GooglePeopleService);
                groupsResource.Delete(group.ResourceName);
            }

            oc.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestSyncDeletedOutlookContactGroup ***");
        }

        [Test]
        public void TestResetMatches()
        {
            Log.Information($"*** TestResetMatches ***");

            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            var oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();

            var gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            var match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.MatchContacts();
            ClassicAssert.IsNotNull(sync.GoogleContacts);
            ClassicAssert.AreEqual(initialGoogleContactCount, sync.GoogleContacts.Count);
            ClassicAssert.IsNotNull(sync.OutlookContacts);
            ClassicAssert.AreEqual(1, sync.OutlookContacts.Count);
            ClassicAssert.IsNotNull(sync.Contacts);
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.Contacts.Count);

            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNotNull(match.GoogleContact);
            ClassicAssert.IsNotNull(match.OutlookContact);
            ContactsMatcher.SyncContact(match, sync);

            // delete outlook contact
            oc.Delete();

            //load the same contact from google
            sync.MatchContacts();
            ClassicAssert.IsNotNull(sync.GoogleContacts);
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.GoogleContacts.Count);
            ClassicAssert.IsNotNull(sync.OutlookContacts);
            ClassicAssert.AreEqual(0, sync.OutlookContacts.Count);
            ClassicAssert.IsNotNull(sync.Contacts);
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.Contacts.Count);

            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNull(match.OutlookContact);
            ClassicAssert.IsNotNull(match.GoogleContact);
            ContactsMatcher.SyncContact(match, sync);

            // reset matches
            var etag = match.GoogleContact.ETag;
            sync.ResetMatch(match.GoogleContact);
            Assert.That(EnsureGoogleContactUpdated(match.GoogleContact, etag));

            // load same contact match
            sync.MatchContacts();
            ClassicAssert.IsNotNull(sync.GoogleContacts);
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.GoogleContacts.Count);
            ClassicAssert.IsNotNull(sync.OutlookContacts);
            ClassicAssert.AreEqual(0, sync.OutlookContacts.Count);
            ClassicAssert.IsNotNull(sync.Contacts);
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.Contacts.Count);
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNull(match.OutlookContact);
            ClassicAssert.IsNotNull(match.GoogleContact);

            //sync.DeleteGoogleResolution = DeleteResolution.KeepGoogleAlways;
            ContactsMatcher.SyncContact(match, sync);

            ClassicAssert.IsNotNull(match.GoogleContact, "google contact should still be present");
            ClassicAssert.IsNotNull(match.OutlookContact, "OutlookContact should be recreated");

            DeleteTestContacts(match);
            Assert.That(EnsureGoogleContactDeleted(match.GoogleContact));

            // create new contact to sync
            oc = new OutlookContactBuilder().Build();
            oc.FullName = TEST_CONTACT_NAME;
            oc.FileAs = TEST_CONTACT_NAME;
            oc.Email1Address = TEST_CONTACT_EMAIL;
            oc.Email2Address = TEST_CONTACT_EMAIL.Replace("00", "01");
            oc.Email3Address = TEST_CONTACT_EMAIL.Replace("00", "02");
            oc.HomeAddress = "10 Parades";
            oc.PrimaryTelephoneNumber = "123";
            oc.Save();

            // same test for delete google contact...
            gc = new GoogleContactBuilder().Build();
            sync.UpdateContact(oc, gc);
            match = new ContactMatch(new OutlookContactInfo(oc, sync), gc);
            match.GoogleContactDirty = true;

            //save contact to google.
            sync.SaveContact(match);
            Assert.That(EnsureGoogleContactSaved(match.GoogleContact));

            //load the same contact from google.
            sync.MatchContacts();
            ClassicAssert.IsNotNull(sync.GoogleContacts);
            ClassicAssert.AreEqual(initialGoogleContactCount, sync.GoogleContacts.Count);
            ClassicAssert.IsNotNull(sync.OutlookContacts);
            ClassicAssert.AreEqual(1, sync.OutlookContacts.Count);
            ClassicAssert.IsNotNull(sync.Contacts);
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.Contacts.Count);

            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNotNull(match.OutlookContact);
            ClassicAssert.IsNotNull(match.GoogleContact);
            ContactsMatcher.SyncContact(match, sync);
            ClassicAssert.IsNotNull(match.OutlookContact);
            ClassicAssert.IsNotNull(match.GoogleContact);

            // delete google contact           
            sync.GooglePeopleResource.DeleteContact(match.GoogleContact.ResourceName).Execute();
            Assert.That(EnsureGoogleContactDeleted(match.GoogleContact));

            //load the same contact from google.
            sync.MatchContacts();
            ClassicAssert.IsNotNull(sync.GoogleContacts);
            ClassicAssert.AreEqual(initialGoogleContactCount, sync.GoogleContacts.Count);
            ClassicAssert.IsNotNull(sync.OutlookContacts);
            ClassicAssert.AreEqual(1, sync.OutlookContacts.Count);
            ClassicAssert.IsNotNull(sync.Contacts);
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.Contacts.Count);

            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ContactsMatcher.SyncContact(match, sync);
            ClassicAssert.IsNotNull(match.OutlookContact);
            ClassicAssert.IsNull(match.GoogleContact);

            // reset matches
            //Not, because null: sync.ResetMatch(match.GoogleContact);
            sync.ResetMatch(match.OutlookContact.GetOriginalItemFromOutlook());

            // load same contact match
            sync.MatchContacts();
            ClassicAssert.IsNotNull(sync.GoogleContacts);
            ClassicAssert.AreEqual(initialGoogleContactCount, sync.GoogleContacts.Count);
            ClassicAssert.IsNotNull(sync.OutlookContacts);
            ClassicAssert.AreEqual(1, sync.OutlookContacts.Count);
            ClassicAssert.IsNotNull(sync.Contacts);
            ClassicAssert.AreEqual(initialGoogleContactCount + 1, sync.Contacts.Count);
            match = sync.ContactByProperty(TEST_CONTACT_NAME, TEST_CONTACT_EMAIL);
            ClassicAssert.IsNotNull(match.OutlookContact);
            ClassicAssert.IsNull(match.GoogleContact);
            //sync.DeleteOutlookResolution = DeleteResolution.DeleteOutlookAlways;
            ContactsMatcher.SyncContact(match, sync);

            // Outlook contact should still be present and GooglePerson should be filled
            ClassicAssert.IsNotNull(match.OutlookContact);
            ClassicAssert.IsNotNull(match.GoogleContact);

            oc.Delete();

            DeleteTestContact(match.GoogleContact);

            Log.Information($"*** TestResetMatches ***");
        }

        private void DeleteTestContacts(ContactMatch match)
        {
            if (match != null)
            {
                DeleteTestContact(match.GoogleContact);
                Assert.That(EnsureGoogleContactDeleted(match.GoogleContact));
                DeleteTestContact(match.OutlookContact);
            }
        }

        private void DeleteTestContact(Outlook.ContactItem oc)
        {
            if (oc != null)
            {
                var name = oc.FileAs;
                oc.Delete();
                Log.Information($"Deleted Outlook test contact: {name}");
            }
        }

        private void DeleteTestContact(OutlookContactInfo oc)
        {
            if (oc != null)
            {
                DeleteTestContact(oc.GetOriginalItemFromOutlook());
            }
        }

        private void DeleteTestContact(Person gc1)
        {
            if (gc1 != null && gc1.Metadata != null && !(gc1.Metadata.Deleted ?? false) && !string.IsNullOrEmpty(gc1.ResourceName))
            {
                try
                {
                    sync.GooglePeopleResource.DeleteContact(gc1.ResourceName).Execute();
                    var googleUniqueIdentifierName = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc1);
                    Log.Information($"Deleted Google test contact: {googleUniqueIdentifierName}");
                }
                catch (Google.GoogleApiException)
                {
                    try
                    {
                        var request = sync.GooglePeopleResource.Get(gc1.ResourceName);
                        request.PersonFields = Synchronizer.GET_PERSON_FIELDS;
                        var gc2 = request.Execute();
                        if (gc2 != null && !(gc2.Metadata.Deleted ?? false))
                        {
                            sync.GooglePeopleResource.DeleteContact(gc2.ResourceName).Execute();
                            Log.Information($"Deleted Google test contact: {ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc1)}");
                        }
                    }
                    catch (Exception e1)
                    {
                        Log.Information(e1, "Exception");
                    }
                }
                catch (Exception e2)
                {
                    Log.Information(e2, "Exception");
                }
            }
        }

        private void DeleteTestContactGroup(ContactGroup g)
        {
            if (g != null && g.Metadata != null && !(g.Metadata.Deleted ?? false))
            {
                try
                {
                    var groupsResource = new ContactGroupsResource(sync.GooglePeopleService);
                    groupsResource.Delete(g.ResourceName).Execute();
                    Log.Information($"Deleted Google test group: {g.Name}");
                }
                catch (Exception e1)
                {
                    Log.Information(e1, "Exception");
                }
            }
        }

        internal ContactMatch FindMatch(Outlook.ContactItem oc)
        {
            foreach (var match in sync.Contacts)
            {
                if (match.OutlookContact != null && match.OutlookContact.EntryID == oc.EntryID)
                {
                    return match;
                }
            }
            return null;
        }

        private void DeleteGoogleTestContacts()
        {
            foreach (var gc in sync.GoogleContacts)
            {
                var pm = ContactPropertiesUtils.GetGooglePrimaryEmailValue(gc);
                var unstructuredName = ContactPropertiesUtils.GetGoogleUnstructuredName(gc);
                var googleFileAs = ContactPropertiesUtils.GetGoogleFileAsValue(gc);
                var org = ContactPropertiesUtils.GetGooglePrimaryOrganizationName(gc);

                if (gc != null &&
                    ((pm == TEST_CONTACT_EMAIL) ||
                      googleFileAs == TEST_CONTACT_NAME ||
                      unstructuredName == TEST_CONTACT_NAME ||
                      org == TEST_CONTACT_NAME))
                {
                    DeleteTestContact(gc);
                    Assert.That(EnsureGoogleContactDeleted(gc));
                }
            }
        }

        private void DeleteOutlookTestContacts()
        {
            var oc = sync.OutlookContacts.Find($"[Email1Address] = '{TEST_CONTACT_EMAIL}'") as Outlook.ContactItem;
            while (oc != null)
            {
                DeleteTestContact(oc);
                oc = sync.OutlookContacts.Find($"[Email1Address] = '{TEST_CONTACT_EMAIL}'") as Outlook.ContactItem;
            }

            oc = sync.OutlookContacts.Find($"[FileAs] = '{TEST_CONTACT_NAME}'") as Outlook.ContactItem;
            while (oc != null)
            {
                DeleteTestContact(oc);
                oc = sync.OutlookContacts.Find($"[FileAs] = '{TEST_CONTACT_NAME}'") as Outlook.ContactItem;
            }

            oc = sync.OutlookContacts.Find($"[FileAs] = '{TEST_SAVE_AS}'") as Outlook.ContactItem;
            while (oc != null)
            {
                DeleteTestContact(oc);
                oc = sync.OutlookContacts.Find($"[FileAs] = '{TEST_SAVE_AS}'") as Outlook.ContactItem;
            }
        }

        private void DeleteTestContacts()
        {
            sync.LoadContacts();
            initialGoogleContactCount = sync.GoogleContacts.Count;

            DeleteOutlookTestContacts();

            //for (var i = 1; i < defaultWaitTries; i++)
            //{
            DeleteGoogleTestContacts();
            //Assert.That(EnsureAllGoogleContactsDeleted())
            //    break;                
            //}

            sync.LoadContacts();
            initialGoogleContactCount = sync.GoogleContacts.Count;
            // ClassicAssert.AreEqual(0, sync.GoogleContacts.Count);
            ClassicAssert.AreEqual(0, sync.OutlookContacts.Count);

            //To be on the safe side: Call again
            //DeleteGoogleTestContacts();
            //sync.LoadContacts();

            initialGoogleContactCount = sync.GoogleContacts.Count;

        }

        private bool EnsureGoogleContactHasPhoto(Person gc)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                //var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"))
                //{
                //    NumberToRetrieve = 256,
                //    StartIndex = 0,
                //    ContactGroup = defaultContactGroup.Id
                //};

                var id = gc.ResourceName;
                //sync.GooglePeopleRequest.PageToken = null;                

                //do
                //{
                //    var response = sync.GooglePeopleRequest.Execute();
                //    foreach (var a in response.Connections)
                //    {
                //        if (id.Equals(a.ResourceName))
                //        {
                //            if (Utilities.HasContactPhoto(a))
                //            {
                //                return true;
                //            }
                //        }
                //    }
                //    sync.GooglePeopleRequest.PageToken = response.NextPageToken;
                //} while (!string.IsNullOrEmpty(sync.GooglePeopleRequest.PageToken));
                var request = sync.GooglePeopleResource.Get(id);
                request.PersonFields = Synchronizer.GET_PERSON_FIELDS;
                var a = request.Execute();
                if (a != null && Utilities.HasContactPhoto(a))
                    return true;

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Log.Information($"EnsureGoogleContactHasPhoto: sleeping for {t}ms");
                Thread.Sleep(t);
            }

            return false;
        }

        private bool EnsureGoogleContactHasGroups(Person gc, int groups)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                var contact = RetrievePerson(ContactPropertiesUtils.GetGoogleId(gc));

                if (contact != null)
                {
                    if (CountGoogleGroups(contact) == groups)
                    {
                        return true;
                    }
                    else
                    {
                        Log.Information($"Found {contact.Memberships.Count} groups");
                    }
                }

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Log.Information($"EnsureGoogleContactHasContactGroups: sleeping for {t}ms");
                Thread.Sleep(t);
            }

            return false;
        }

        private bool EnsureGoogleGroupExist(string gn)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (RetrieveContactGroup(gn))
                {
                    return true;
                }

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Log.Information($"EnsureGoogleContactGroupSaved: sleeping for {t}ms");
                Thread.Sleep(t);
            }
            return false;
        }

        private bool EnsureGoogleContactSaved(Person gc)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (RetrievePerson(ContactPropertiesUtils.GetGoogleId(gc)) != null)
                {
                    return true;
                }

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Log.Information($"EnsureGoogleContactSaved: sleeping for {t}ms");
                Thread.Sleep(t);
            }
            return false;
        }

        private bool EnsureGoogleContactDeleted(Person gc)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (RetrievePerson(ContactPropertiesUtils.GetGoogleId(gc)) == null)
                {
                    return true;
                }

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Log.Information($"EnsureGoogleContactDeleted: sleeping for {t}ms");
                Thread.Sleep(t);
            }
            return false;
        }

        //private bool EnsureAllGoogleContactsDeleted()
        //{
        //    foreach (var gc in sync.GoogleContacts)
        //    {
        //        //for (var i = 1; i < defaultWaitTries; i++)
        //        //{
        //            if (RetrievePerson(gc.ResourceName) != null)
        //            {
        //                return false;
        //            }

        //            //var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
        //            //Log.Information($"EnsureGoogleContactUpdated: sleeping for {t}ms");
        //            //Thread.Sleep(t);
        //        //}
        //    }

        //    return true;
        //}

        private bool EnsureGoogleContactUpdated(Person gc, string etag)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (RetrieveContactIfUpdated(gc, etag))
                {
                    return true;
                }

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Log.Information($"EnsureGoogleContactUpdated: sleeping for {t}ms");
                Thread.Sleep(t);
            }
            return false;
        }

        private bool EnsureGoogleGroupsCount(int count)
        {
            for (var i = 1; i < defaultWaitTries; i++)
            {
                if (CountGoogleGroups() == count)
                {
                    return true;
                }

                var t = (int)(Math.Pow(2.0, i - 1) * defaultWait);
                Log.Information($"EnsureGoogleGroupsCount: sleeping for {t}ms");
                Thread.Sleep(t);
            }
            return false;
        }

        private bool RetrieveContactIfUpdated(Person gc, string etag)
        {
            //var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"))
            //{
            //    NumberToRetrieve = 256,
            //    StartIndex = 0,
            //    ContactGroup = defaultContactGroup.Id
            //};
            var id = gc.ResourceName;

            //ToDo:Unfortunately not the same as the Alternative below, because same Resource is returned, but different eTag
            //var request = sync.GooglePeopleResource.Get(id);
            //request.PersonFields = Synchronizer.GET_PERSON_FIELDS;
            //var a = request.Execute();
            //if (a != null && a.ETag != etag)
            //    return true;

            //Alternative, safer but also slower: Load all Google Contacts and go thru to check eTag
            sync.GooglePeopleRequest.PageToken = null;

            do
            {
                var response = sync.GooglePeopleRequest.Execute();
                if (response.Connections != null)
                {
                    foreach (var a in response.Connections)
                    {
                        if (id.Equals(a.ResourceName))
                            if (a.ETag != etag)
                                return true;
                    }
                }
                sync.GooglePeopleRequest.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(sync.GooglePeopleRequest.PageToken));
            return false;
        }

        //private bool RetrievePerson(Person gc)
        //{
        //    //var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"))
        //    //{
        //    //    NumberToRetrieve = 256,
        //    //    StartIndex = 0,
        //    //    ContactGroup = defaultContactGroup.Id
        //    //};

        //    var id = gc.ResourceName;
        //    var feed = sync.GooglePeopleResource.Connections.List("people/me").Execute();

        //    //while (feed != null)
        //    //{
        //        foreach (var a in feed.Connections)
        //        {
        //            if (id.Equals(a.ResourceName))
        //            {
        //                return true;
        //            }
        //        }
        //        //query.StartIndex += query.NumberToRetrieve;
        //        //feed = sync.ContactsRequest.Get(feed, FeedRequestType.Next);
        //    //}
        //    return false;
        //}

        private Person RetrievePerson(string id)
        {
            var request = sync.GooglePeopleRequest;
            request.PersonFields = Synchronizer.GET_PERSON_FIELDS;

            request.PageToken = null;

            do
            {
                var response = request.Execute();
                if (response.Connections != null)
                {
                    foreach (var a in response.Connections)
                    {
                        if (a.Metadata != null && !(a.Metadata.Deleted ?? false))
                        {
                            var a_id = ContactPropertiesUtils.GetGoogleId(a);
                            if (!string.IsNullOrEmpty(id) && id.Equals(a_id, StringComparison.InvariantCultureIgnoreCase))
                            {
                                return a;

                            }
                        }
                    }
                }
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));

            return null;

        }

        /*private Person RetrievePerson(string resourceName)
        {
            try
            {
                //var u = new Uri(gid.Replace("http://", "https://"));
                var request = sync.GooglePeopleResource.Get(resourceName);
                request.PersonFields = Synchronizer.GET_PERSON_FIELDS;
                var response = request.Execute();                
                if (response == null || (response.Metadata != null && (response.Metadata.Deleted??false)))
                    return null;

                return response;
            }
            catch (Exception ex)
            {
                Log.Information(ex, "Exception");
                return null;
            }
        }*/

        private bool RetrieveContactGroup(string gn)
        {
            //var query = new ContactGroupsQuery(ContactGroupsQuery.CreateContactGroupsUri("default"))
            //{
            //    NumberToRetrieve = 256,
            //    StartIndex = 0
            //};

            var groupsResource = new ContactGroupsResource(sync.GooglePeopleService);
            var groupsRequest = groupsResource.List();
            groupsRequest.PageToken = null;

            do
            {
                var response = groupsRequest.Execute();
                if (response.ContactGroups != null)
                {
                    foreach (var a in response.ContactGroups)
                    {
                        if (a.Name == gn)
                        {
                            return true;
                        }
                    }
                }
                groupsRequest.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(groupsRequest.PageToken));
            return false;
        }

        private void DeleteTestContactGroups()
        {
            //var query = new ContactGroupsQuery(ContactGroupsQuery.CreateContactGroupsUri("default"))
            //{
            //    NumberToRetrieve = 256,
            //    StartIndex = 0
            //};

            var groupsResource = new ContactGroupsResource(sync.GooglePeopleService);
            var groupsRequest = groupsResource.List();
            groupsRequest.PageToken = null;

            do
            {
                var response = groupsRequest.Execute();
                if (response.ContactGroups != null)
                {
                    foreach (var a in response.ContactGroups)
                    {
                        if (IsTestContactGroup(a))
                        {
                            DeleteTestContactGroup(a);
                        }
                    }
                }

                groupsRequest.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(groupsRequest.PageToken));
        }

        private int CountGoogleGroups()
        {
            //var query = new ContactGroupsQuery(ContactGroupsQuery.CreateContactGroupsUri("default"))
            //{
            //    NumberToRetrieve = 256,
            //    StartIndex = 0
            //};

            var count = 0;

            var groupsResource = new ContactGroupsResource(sync.GooglePeopleService);
            var groupsRequest = groupsResource.List();
            groupsRequest.PageToken = null;

            do
            {
                var response = groupsRequest.Execute();
                if (response.ContactGroups != null)
                {
                    foreach (var a in response.ContactGroups)
                    {
                        if (a.Metadata != null && !(a.Metadata.Deleted ?? false) && a.ResourceName != Synchronizer.myContactsGroup)
                        {
                            count++;
                        }
                    }
                }
                groupsRequest.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(groupsRequest.PageToken));

            return count;
        }

        private int CountGoogleGroups(Person a)
        {
            var ret = 0;
            if (a.Memberships != null)
                foreach (var group in a.Memberships)
                    if (group.ContactGroupMembership.ContactGroupResourceName != Synchronizer.myContactsGroup)
                        ret++;

            return ret;
        }

        private void DeleteAppointmentsForTestContacts()
        {
            //Also delete the birthday/anniversary entries in Outlook calendar
            Log.Information("Deleting Outlook calendar TEST entries (birthday, anniversary) ...");

            try
            {
                var outlookNamespace = Synchronizer.OutlookApplication.GetNamespace("mapi");
                var calendarFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                var outlookCalendarItems = calendarFolder.Items;
                for (var i = outlookCalendarItems.Count; i > 0; i--)
                {
                    var item = outlookCalendarItems[i] as Outlook.AppointmentItem;
                    if (item.Subject.Contains(TEST_CONTACT_NAME))
                    {
                        var subject = item.Subject;
                        item.Delete();
                        Log.Information($"Deleted Outlook calendar TEST entry: {subject}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Information($"Could not delete Outlook calendar TEST entries: {ex.Message}");
            }
        }

        private bool IsTestContactGroup(ContactGroup g)
        {
            return g.Name == TEST_GROUP;
        }
        private void DelayBetweenSync()
        {
            //we need to wait at least 3 minutes, so next synchronization will not ignore updates, or are 2 minutes enough?
            Thread.Sleep(2 *60 * 1000 +1);
        }

        private static Microsoft.Win32.RegistryKey LoadSettings(out string gmailUsername, out string syncProfile)
        {
            //sync.LoginToGoogle(ConfigurationManager.AppSettings["Gmail.Username"],  ConfigurationManager.AppSettings["Gmail.Password"]);
            //ToDo: Reading the username and config from the App.Config file doesn't work. If it works, consider special characters like & = &amp; in the XML file
            //ToDo: Maybe add a common Test account to the App.config and read it from there, encrypt the password
            //For now, read the userName from the Registry (same settings as for GoogleContactsSync Application
            gmailUsername = "";

            const string appRootKey = SettingsForm.AppRootKey;
            var regKeyAppRoot = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRootKey);
            syncProfile = "Default Profile";
            if (regKeyAppRoot.GetValue("SyncProfile") != null)
            {
                syncProfile = regKeyAppRoot.GetValue("SyncProfile") as string;
            }

            regKeyAppRoot = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRootKey + (syncProfile != null ? ('\\' + syncProfile) : ""));

            if (regKeyAppRoot.GetValue("Username") != null)
            {
                gmailUsername = regKeyAppRoot.GetValue("Username") as string;

            }

            return regKeyAppRoot;
        }

    }
}