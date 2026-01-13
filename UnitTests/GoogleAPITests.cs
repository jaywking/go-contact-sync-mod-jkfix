using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Util.Store;
using Event = Google.Apis.Calendar.v3.Data.Event;
using NodaTime;
using NUnit.Framework;
using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Net;
using Polly.Contrib.WaitAndRetry;
using Polly;
using Polly.Registry;
using Polly.Retry;
using Polly.Wrap;
using NUnit.Framework.Legacy;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class GoogleAPITests
    {
        private EventsResource eventsService = null;
        private CalendarListEntry primaryCalendar = null;

        [OneTimeSetUp]
        public void Init()
        {
            LoadSettings(out var gmailUsername, out _);

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
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(stream).Secrets, scopes, gmailUsername, CancellationToken.None,
                fDS).Result;

                var initializer = new Google.Apis.Services.BaseClientService.Initializer
                {
                    HttpClientInitializer = credential
                };
                var CalendarRequest = GoogleServices.CreateCalendarService(initializer);

                var list = CalendarRequest.CalendarList.List().Execute().Items;
                foreach (var calendar in list)
                {
                    if (calendar.Primary != null && calendar.Primary.Value)
                    {
                        primaryCalendar = calendar;
                        break;
                    }
                }

                if (primaryCalendar == null)
                {
                    throw new Exception("Primary Calendar not found");
                }
                eventsService = CalendarRequest.Events;
            }
        }

        [Test]
        public void CreateNewPerson()
        {
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

                //var parameters = new OAuth2Parameters
                //{
                //    ClientId = clientSecrets.Secrets.ClientId,
                //    ClientSecret = clientSecrets.Secrets.ClientSecret,

                //    // Note: AccessToken is valid only for 60 minutes
                //    AccessToken = credential.Token.AccessToken,
                //    RefreshToken = credential.Token.RefreshToken
                //};

                //var settings = new RequestSettings("GoContactSyncMod", parameters);

                service = GoogleServices.CreatePeopleService(initializer);
            }

            #region Delete previously created test contact.
            var peopleResource = new PeopleResource(service);
            var peopleRequest = peopleResource.Connections.List("people/me");
            peopleRequest.PersonFields = Synchronizer.GET_PERSON_FIELDS;
            peopleRequest.PageToken = null;

            //Log.Information("Loaded Google contacts");

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
                            //break;
                        }

                        peopleRequest.PageToken = response.NextPageToken;
                    }
                }
            } while (!string.IsNullOrEmpty(peopleRequest.PageToken));
            #endregion

            var newEntry = new GoogleContactBuilder().Build("Who is this guy?");

            //var feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

            var createdEntry = peopleResource.CreateContact(newEntry).Execute();

            Log.Information("Created Google contact");

            ClassicAssert.IsNotNull(createdEntry.ResourceName);
            var updateRequest = peopleResource.UpdateContact(createdEntry, createdEntry.ResourceName);
            updateRequest.UpdatePersonFields = Synchronizer.UPDATE_PERSON_FIELDS;
            _ = updateRequest.Execute();

            Log.Information("Updated Google contact");

            //delete test contacts
            peopleResource.DeleteContact(createdEntry.ResourceName).Execute();

            Log.Information("Deleted Google contact");
        }


        [Test]
        public void CreateNewAppointment()
        {
            #region Delete previously created test contact.
            var query = eventsService.List(primaryCalendar.Id);
            query.MaxResults = 500;
            query.TimeMinDateTimeOffset = DateTime.Now.AddDays(-10);
            query.TimeMaxDateTimeOffset = DateTime.Now.AddDays(10);

            var feed = query.Execute();
            Log.Information("Loaded Google appointments");
            foreach (var entry in feed.Items)
            {
                if (entry.Summary != null && entry.Summary.Contains("GCSM Test Appointment") && !entry.Status.Equals("cancelled"))
                {
                    Log.Information("Deleting Google appointment:" + entry.Summary + " - " + entry.Start.DateTimeDateTimeOffset.ToString());
                    eventsService.Delete(primaryCalendar.Id, entry.Id);
                    Log.Information("Deleted Google appointment");
                    //break;
                }
            }
            #endregion

            var newEntry = Factory.NewEvent();
            newEntry.Summary = "GCSM Test Appointment";
            newEntry.Start.DateTimeDateTimeOffset = new DateTimeOffset(DateTime.Now);
            newEntry.End.DateTimeDateTimeOffset = new DateTimeOffset(DateTime.Now);

            var createdEntry = eventsService.Insert(newEntry, primaryCalendar.Id).Execute();

            Log.Information("Created Google appointment");

            ClassicAssert.IsNotNull(createdEntry.Id);

            var updatedEntry = eventsService.Update(createdEntry, primaryCalendar.Id, createdEntry.Id).Execute();

            Log.Information("Updated Google appointment");

            //delete test contacts
            eventsService.Delete(primaryCalendar.Id, updatedEntry.Id).Execute();

            Log.Information("Deleted Google appointment");
        }

        [Test]
        [Ignore("Use only when it is needed as it overloads Google People Read API quota")]
        public void TestPeopleReadApiQuota()
        {
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

            //Log.Information("Loaded Google contacts");

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
                            //break;
                        }

                        peopleRequest.PageToken = response.NextPageToken;
                    }
                }
            } while (!string.IsNullOrEmpty(peopleRequest.PageToken));
            #endregion

            var newEntry = new GoogleContactBuilder().Build("Quota Test");

            var createdEntry = peopleResource.CreateContact(newEntry).Execute();

            Log.Information("Created Google contact");

            ClassicAssert.IsNotNull(createdEntry.ResourceName);

            var delay = Backoff.ExponentialBackoff(TimeSpan.FromMilliseconds(1000), retryCount: 5);

            var policyContactRead = Policy
                .Handle<Google.GoogleApiException>(ex => ex.HttpStatusCode == (HttpStatusCode)429 && ex.Error.Message.StartsWith("Quota exceeded for quota metric"))
                .WaitAndRetry(delay, onRetry: (exception, retryCount, context) =>
                {
                    Log.Debug($"Retry, waiting for {retryCount}");
                });

            var registry = new PolicyRegistry()
            {
                { "Contact Read", policyContactRead }
            };

            try
            {
                for (var i = 0; i < 1000; i++)
                {
                    do
                    {
                        var policy = registry.Get<Policy>("Contact Read");

                        ListConnectionsResponse response = null;
                        var result = policy.ExecuteAndCapture(() =>
                        {
                            response = peopleRequest.Execute();
                        });

                        if (response?.Connections != null)
                        {
                            foreach (var entry in response.Connections)
                            {
                                var pm = ContactPropertiesUtils.GetGooglePrimaryEmail(entry);
                                if (pm != null && pm.Value == SyncContactsTests.TEST_CONTACT_EMAIL)
                                {
                                    Log.Information("Read Google contact");
                                }

                                peopleRequest.PageToken = response.NextPageToken;
                            }
                        }
                    } while (!string.IsNullOrEmpty(peopleRequest.PageToken));
                }
            }
            catch (Google.GoogleApiException ex)
            {
                if (ex.HttpStatusCode == (HttpStatusCode)429 && ex.Error.Message.StartsWith("Quota exceeded for quota metric"))
                {
                    Log.Information(ex, "Exception");
                }
            }
            catch (Exception ex)
            {
                Log.Information(ex, "Exception");
            }

            //delete test contacts
            peopleResource.DeleteContact(createdEntry.ResourceName).Execute();

            Log.Information("Deleted Google contact");
        }


        [Test]
        [Ignore("Use only when it is needed as it overloads Google People Write API quota")]
        public void TestPeopleWriteApiQuota()
        {
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

            //Log.Information("Loaded Google contacts");

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
                            //break;
                        }

                        peopleRequest.PageToken = response.NextPageToken;
                    }
                }
            } while (!string.IsNullOrEmpty(peopleRequest.PageToken));
            #endregion

            var newEntry = new GoogleContactBuilder().Build("Quota Test");

            //var feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

            var createdEntry = peopleResource.CreateContact(newEntry).Execute();

            Log.Information("Created Google contact");

            ClassicAssert.IsNotNull(createdEntry.ResourceName);

            var delay = Backoff.ExponentialBackoff(TimeSpan.FromMilliseconds(1000), retryCount: 5);

            var policyContactWrite1 = Policy
                .Handle<Google.GoogleApiException>(ex => ex.HttpStatusCode == (HttpStatusCode)429 && ex.Error.Message.StartsWith("Quota exceeded for quota metric"))
                .WaitAndRetry(delay, onRetry: (exception, retryCount, context) =>
                {
                    Log.Debug($"Retry, waiting for {retryCount}");
                });

            var policyProtocolViolationException1 = Policy
                .Handle<ProtocolViolationException>()
                .Retry(1, onRetry: (exception, retryCount) =>
                {
                    Log.Debug($"Retry {retryCount}");
                });

            var registry = new PolicyRegistry()
                    {
                           { "Contact Write", policyContactWrite1 },
                           { "ProtocolViolationException", policyProtocolViolationException1 }
                    };

            var policyWrap1 = Policy.Wrap(policyProtocolViolationException1, policyContactWrite1);


            var registry1 = new PolicyRegistry()
            {
                           { "Contact Write", policyWrap1 }
            };

            try
            {
                for (var i = 0; i < 1000; i++)
                {
                    var updateRequest = peopleResource.UpdateContact(createdEntry, createdEntry.ResourceName);
                    updateRequest.UpdatePersonFields = Synchronizer.UPDATE_PERSON_FIELDS;

                    var policyContactWrite = registry.Get<RetryPolicy>("Contact Write");
                    var policyProtocolViolationException = registry.Get<RetryPolicy>("ProtocolViolationException");

                    //var policyWrap = Policy.Wrap(policyProtocolViolationException, policyContactWrite);

                    var policyWrap = registry1.Get<PolicyWrap>("Contact Write");

                    var result = policyWrap.ExecuteAndCapture(() =>
                    {
                        //throw new ProtocolViolationException();
                        createdEntry = updateRequest.Execute();
                    });
                }
            }
            catch (Google.GoogleApiException ex)
            {
                if (ex.HttpStatusCode == (HttpStatusCode)429 && ex.Error.Message.StartsWith("Quota exceeded for quota metric"))
                {
                    Log.Information(ex, "Exception");
                }
            }
            catch (Exception ex)
            {

                Log.Information(ex, "Exception");

            }
            Log.Information("Updated Google contact");

            //delete test contacts
            peopleResource.DeleteContact(createdEntry.ResourceName).Execute();

            Log.Information("Deleted Google contact");
        }

        [Test]
        public void Test_OldRecurringAppointment()
        {
            #region Delete previously created test contact.
            var query = eventsService.List(primaryCalendar.Id);
            query.MaxResults = 500;
            query.TimeMinDateTimeOffset = DateTime.Now.AddDays(-10);
            query.TimeMaxDateTimeOffset = DateTime.Now.AddDays(10);
            //query.Q = "GCSM Test Appointment";

            var feed = query.Execute();
            Log.Information("Loaded Google appointments");
            foreach (var entry in feed.Items)
            {
                if (entry.Summary != null && entry.Summary.Contains("GCSM Test Appointment") && !entry.Status.Equals("cancelled"))
                {
                    Log.Information("Deleting Google appointment:" + entry.Summary + " - " + entry.Start.DateTimeDateTimeOffset.ToString());
                    eventsService.Delete(primaryCalendar.Id, entry.Id);
                    Log.Information("Deleted Google appointment");
                    //break;
                }
            }

            #endregion

            var zone = DateTimeZoneProviders.Tzdb["Europe/Warsaw"];

            var e1_start = new LocalDateTime(1970, 10, 14, 10, 0, 0);
            var e1_start_zoned = e1_start.InZoneLeniently(zone);
            var e1_start_utc = e1_start_zoned.ToDateTimeUtc();
            _ = new LocalDateTime(1970, 10, 14, 11, 0, 0);
            _ = e1_start.InZoneLeniently(zone);
            var e1_end_utc = e1_start_zoned.ToDateTimeUtc();

            var s = new EventDateTime
            {
                DateTimeDateTimeOffset = new DateTimeOffset(e1_start_utc),
                TimeZone = "Europe/Warsaw"
            };

            var e = new EventDateTime
            {
                DateTimeDateTimeOffset = new DateTimeOffset(e1_end_utc),
                TimeZone = "Europe/Warsaw"
            };

            var e1 = new Event()
            {
                Summary = "Birthday 1",
                Start = s,
                End = e,
                Recurrence = new string[] { "RRULE:FREQ=YEARLY;BYMONTHDAY=14;BYMONTH=10" }
            };

            ClassicAssert.AreEqual("1970-10-14T09:00:00Z", e1.Start.DateTimeRaw);
            var c1 = eventsService.Insert(e1, primaryCalendar.Id).Execute();
            ClassicAssert.AreEqual("1970-10-14T10:00:00+01:00", c1.Start.DateTimeRaw);

            var e2_start = new LocalDateTime(2000, 10, 14, 10, 0, 0);
            var e2_start_zoned = e2_start.InZoneLeniently(zone);
            var e2_start_utc = e2_start_zoned.ToDateTimeUtc();
            _ = new LocalDateTime(2000, 10, 14, 11, 0, 0);
            _ = e2_start.InZoneLeniently(zone);
            var e2_end_utc = e2_start_zoned.ToDateTimeUtc();

            var ss = new EventDateTime
            {
                DateTimeDateTimeOffset = new DateTimeOffset(e2_start_utc),
                TimeZone = "Europe/Warsaw"
            };

            var ee = new EventDateTime
            {
                DateTimeDateTimeOffset = new DateTimeOffset(e2_end_utc),
                TimeZone = "Europe/Warsaw"
            };

            var e2 = new Event()
            {
                Summary = "Birthday 2",
                Start = ss,
                End = ee,
                Recurrence = new string[] { "RRULE:FREQ=YEARLY;BYMONTHDAY=14;BYMONTH=10" }
            };

            ClassicAssert.AreEqual("2000-10-14T08:00:00Z", e2.Start.DateTimeRaw);
            var c2 = eventsService.Insert(e2, primaryCalendar.Id).Execute();
            ClassicAssert.AreEqual("2000-10-14T10:00:00+02:00", c2.Start.DateTimeRaw);

            Log.Information("Created Google appointment");

            ClassicAssert.IsNotNull(c1.Id);

            //delete test contacts
            eventsService.Delete(primaryCalendar.Id, c1.Id).Execute();
            eventsService.Delete(primaryCalendar.Id, c2.Id).Execute();

            Log.Information("Deleted Google appointment");
        }

        [Test]
        public void Test_RetrieveRecurrenceInstance()
        {
            var e = new Event()
            {
                Summary = "AN_OUTLOOK_TEST_APPOINTMENT",
                Start = new EventDateTime()
                {
                    DateTimeDateTimeOffset = new DateTimeOffset(new DateTime(2020, 6, 3, 15, 0, 0)),
                    TimeZone = "Europe/Warsaw"
                },
                End = new EventDateTime()
                {
                    DateTimeDateTimeOffset = new DateTimeOffset(new DateTime(2020, 6, 3, 16, 0, 0)),
                    TimeZone = "Europe/Warsaw"
                },
                Recurrence = new string[] { "RRULE:FREQ=WEEKLY;UNTIL=20200625;BYDAY=WE" }
            };

            e = eventsService.Insert(e, primaryCalendar.Id).Execute();

            ClassicAssert.IsNotNull(e.Id);

            var r = eventsService.Instances(primaryCalendar.Id, e.Id);

            var dt = new DateTime(2020, 6, 10, 13, 0, 0, DateTimeKind.Utc);
            var s = dt.ToLocalTime().ToString("yyyy-MM-ddTHH:mm:sszzz");
            ClassicAssert.AreEqual("2020-06-10T15:00:00+02:00", s);

            r.OriginalStart = "2020-06-10T15:00:00+02:00";
            var instances = r.Execute();
            ClassicAssert.IsNotNull(instances);
            ClassicAssert.AreEqual(1, instances.Items.Count);

            //delete test contacts
            eventsService.Delete(primaryCalendar.Id, e.Id).Execute();

            Log.Information("Deleted Google appointment");
        }

        [Test]
        public void CreateContactWithLargeNotes()
        {
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

                /*var parameters = new OAuth2Parameters
                {
                    ClientId = clientSecrets.Secrets.ClientId,
                    ClientSecret = clientSecrets.Secrets.ClientSecret,

                    // Note: AccessToken is valid only for 60 minutes
                    AccessToken = credential.Token.AccessToken,
                    RefreshToken = credential.Token.RefreshToken
                };

                var settings = new RequestSettings("GoContactSyncMod", parameters);*/

                service = GoogleServices.CreatePeopleService(initializer);
            }

            #region Delete previously created test contact.
            var peopleResource = new PeopleResource(service);
            var peopleRequest = peopleResource.Connections.List("people/me");
            peopleRequest.PersonFields = Synchronizer.GET_PERSON_FIELDS;
            peopleRequest.PageToken = null;


            Log.Information("Loaded Google contacts");

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
                            //break;
                        }
                        peopleRequest.PageToken = response.NextPageToken;
                    }
                }
            } while (!string.IsNullOrEmpty(peopleRequest.PageToken));
            #endregion

            var newEntry = new GoogleContactBuilder().Build(new string('*', 150000));

            //var feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

            try
            {
                var createdEntry = peopleResource.CreateContact(newEntry).Execute();

                Log.Information("Created Google contact");

                ClassicAssert.IsNotNull(createdEntry.ResourceName);
                var updateRequest = peopleResource.UpdateContact(createdEntry, createdEntry.ResourceName);
                updateRequest.UpdatePersonFields = Synchronizer.UPDATE_PERSON_FIELDS;
                _ = updateRequest.Execute();

                Log.Information("Updated Google contact");

                //delete test contacts
                peopleResource.DeleteContact(createdEntry.ResourceName).Execute();
            }
            catch (Google.GoogleApiException ex) when (ex.Error != null && ex.Error.ErrorResponseContent.Contains("Resource has been exhausted (e.g. check quota)")) //Old Google Contacts Error: "Request data is too large."
            {
                Log.Information(ex, "Exception");
            }

            Log.Information("Deleted Google contact");
        }

        internal static void LoadSettings(out string gmailUsername, out string syncProfile, out string syncContactsFolder, out string syncAppointmentsFolder)
        {
            var regKeyAppRoot = LoadSettings(out gmailUsername, out syncProfile);

            syncContactsFolder = "";
            syncAppointmentsFolder = "";
            Synchronizer.SyncAppointmentsGoogleFolder = "";

            //First, check if there is a folder called GCSMTestContacts available, if yes, use them
            var outlookContactFolders = new ArrayList();
            var outlookAppointmentFolders = new ArrayList();

            try
            { //Add Default Contacts Folder
                var defaultFolder = Synchronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
                outlookContactFolders.Add(new OutlookFolder(defaultFolder.FolderPath, defaultFolder.EntryID, true));
            }
            catch (Exception e)
            {
                Log.Debug(e, "Exception");
                Log.Warning("Error adding OlDefaultFolders.olFolderContacts: " + e.Message);
            }

            try
            {//Add Default Calendar/Appointment folder
                var defaultFolder = Synchronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);
                outlookAppointmentFolders.Add(new OutlookFolder(defaultFolder.FolderPath, defaultFolder.EntryID, true));
            }
            catch (Exception e)
            {
                Log.Debug(e, "Exception");
                Log.Warning("Error adding OlDefaultFolders.olFolderContacts: " + e.Message);
            }

            var folders = Synchronizer.OutlookNameSpace.Folders;
            foreach (Microsoft.Office.Interop.Outlook.Folder folder in folders)
            {
                try
                {
                    SettingsForm.GetOutlookMAPIFolders(outlookContactFolders, outlookAppointmentFolders, folder);
                }
                catch (Exception e)
                {
                    Log.Warning("Error getting available Outlook folders: " + e.Message);
                }
            }

            foreach (OutlookFolder folder in outlookContactFolders)
            {
                if (folder.FolderName.ToUpper().Contains("GCSMTestContacts".ToUpper()))
                {
                    Log.Information("Uses Test folder: " + folder.DisplayName);
                    syncContactsFolder = folder.FolderID;
                    break;
                }
            }

            foreach (OutlookFolder folder in outlookAppointmentFolders)
            {
                if (folder.FolderName.ToUpper().Contains("GCSMTestAppointments".ToUpper()))
                {
                    Log.Information("Uses Test folder: " + folder.DisplayName);
                    syncAppointmentsFolder = folder.FolderID;
                    break;
                }
            }

            if (string.IsNullOrEmpty(syncContactsFolder))
            {
                if (regKeyAppRoot.GetValue("SyncContactsFolder") != null)
                {
                    syncContactsFolder = regKeyAppRoot.GetValue("SyncContactsFolder") as string;
                }
            }

            if (string.IsNullOrEmpty(syncAppointmentsFolder))
            {
                if (regKeyAppRoot.GetValue("SyncAppointmentsFolder") != null)
                {
                    syncAppointmentsFolder = regKeyAppRoot.GetValue("SyncAppointmentsFolder") as string;
                }
            }

            if (string.IsNullOrEmpty(Synchronizer.SyncAppointmentsGoogleFolder))
            {
                if (regKeyAppRoot.GetValue("SyncAppointmentsGoogleFolder") != null)
                {
                    Synchronizer.SyncAppointmentsGoogleFolder = regKeyAppRoot.GetValue("SyncAppointmentsGoogeFolder") as string;
                }
            }
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
