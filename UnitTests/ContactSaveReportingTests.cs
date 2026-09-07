using Google.Apis.Http;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    [NonParallelizable]
    public class ContactSaveReportingTests
    {
        private static ContactMatch Contact(bool existing = true)
        {
            // Construct the cached Outlook data only; never open Outlook or an account.
            var outlook = (OutlookContactInfo)Activator.CreateInstance(typeof(OutlookContactInfo), true);
            outlook.FileAs = "Save reporting test";
            outlook.UserProperties = new OutlookContactInfo.UserPropertiesHolder { GoogleContactId = "test" };
            return new ContactMatch(outlook, new Person
            {
                ResourceName = existing ? "people/test" : null,
                Names = new List<Name> { new Name
                {
                    DisplayName = outlook.FileAs, UnstructuredName = outlook.FileAs,
                    Metadata = new FieldMetadata { Primary = true }
                } },
                ClientData = new List<ClientData>
                {
                    new ClientData { Key = OutlookPropertiesUtils.GetKey(), Value = "outlook-test" }
                }
            }) { GoogleContactDirty = true };
        }

        [TestCase(false)]
        [TestCase(true)]
        public void FailedSaveCountsOnceAndDoesNotSubtractPreviousSuccess(bool throws)
        {
            var sync = new Synchronizer();
            var failures = new List<Exception>();
            sync.ErrorEncountered += (title, error) => failures.Add(error);
            var failure = new ApplicationException("Simulated write failure");
            var calls = 0;
            sync.SaveContacts(new List<ContactMatch> { Contact(), Contact(), Contact() },
                match => sync.SaveContact(match, _ =>
                {
                    if (++calls != 2) return true;
                    if (throws) throw failure;
                    return false;
                }));
            Assert.That(calls, Is.EqualTo(3), "Continue to the next contact after reporting a failure.");
            Assert.That(sync.SyncedCount, Is.EqualTo(2));
            Assert.That(sync.ErrorCount, Is.EqualTo(1));
            Assert.That(failures.Count, Is.EqualTo(1));
            if (throws) Assert.That(failures[0].InnerException, Is.SameAs(failure));
        }

        [Test]
        public void FirstContactFailureCannotMakeSyncedCountNegative()
        {
            var sync = new Synchronizer();
            sync.ErrorEncountered += (title, error) => { };
            sync.SaveContacts(new List<ContactMatch> { Contact() },
                match => sync.SaveContact(match, _ => throw new ApplicationException("Failed")));
            Assert.That(sync.ErrorCount, Is.EqualTo(1));
            Assert.That(sync.SyncedCount, Is.Zero);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void WithoutErrorSubscriberFailureIsCountedAndPropagated(bool throws)
        {
            var sync = new Synchronizer();
            var failure = new ApplicationException("Failed");
            var caught = Assert.Throws<ApplicationException>(() => sync.SaveContacts(
                new List<ContactMatch> { Contact() }, _ => throws ? throw failure : false));
            if (throws) Assert.That(caught, Is.SameAs(failure));
            Assert.That(sync.ErrorCount, Is.EqualTo(1));
            Assert.That(sync.SyncedCount, Is.Zero);
        }

        [Test]
        public void UnchangedAndSkippedContactsDoNotCountAsSuccessfulWrites()
        {
            var sync = new Synchronizer { SyncOption = SyncOption.GoogleToOutlookOnly };
            var unchanged = Contact();
            unchanged.GoogleContactDirty = false;
            sync.SaveContacts(new List<ContactMatch>
            {
                unchanged, new ContactMatch(null, Contact().GoogleContact)
            }, match => sync.SaveContact(match, _ => throw new Exception("Unexpected write")));
            Assert.That(sync.ErrorCount, Is.Zero);
            Assert.That(sync.SyncedCount, Is.Zero);
            Assert.That(sync.SkippedCount, Is.EqualTo(1));
        }

        [TestCase(false, "Permission denied")]
        [TestCase(true, "Permission denied")]
        [TestCase(false, "Resource has been exhausted (e.g. check quota)")]
        [TestCase(true, "Resource has been exhausted (e.g. check quota)")]
        [TestCase(true, "Invalid country code: ZZ")]
        [TestCase(true, "extendedProperty count limit exceeded: 10")]
        public void GoogleFailureReachesSummaryAndKeepsOriginalReason(bool existing, string reason)
        {
            var handler = new FakeHandler(_ => new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent(JsonConvert.SerializeObject(new
                {
                    error = new { code = 400, message = reason }
                }), System.Text.Encoding.UTF8, "application/json")
            });
            using (var service = Service(handler))
            {
                var sync = SynchronizerFor(service);
                Exception reported = null;
                sync.ErrorEncountered += (title, error) => reported = error;
                sync.SaveContacts(new List<ContactMatch> { Contact(existing) },
                    match => sync.SaveContact(match, _ => sync.SaveGoogleContact(match.GoogleContact) != null));
                Assert.That(sync.ErrorCount, Is.EqualTo(1));
                Assert.That(sync.SyncedCount, Is.Zero);
                Assert.That(reported, Is.Not.Null);
                Assert.That(reported.Message, Does.Contain(reason));
                Assert.That(reported.GetBaseException(), Is.TypeOf<Google.GoogleApiException>());
                Assert.That(handler.Calls, Is.EqualTo(1), "Permanent failures must not retry recursively.");
            }
        }

        [TestCase(false, false)]
        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(true, true)]
        public void ProtocolRetryReportsOnlyItsFinalOutcome(bool existing, bool recovers)
        {
            var handler = new FakeHandler(call =>
            {
                if (!recovers || call == 1) throw new ProtocolViolationException("Simulated transport failure");
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("{\"resourceName\":\"people/test\"}", System.Text.Encoding.UTF8, "application/json")
                };
            });
            using (var service = Service(handler))
            {
                var sync = SynchronizerFor(service);
                var errors = 0;
                sync.ErrorEncountered += (title, error) => errors++;
                sync.SaveContacts(new List<ContactMatch> { Contact(existing) },
                    match => sync.SaveContact(match, _ => sync.SaveGoogleContact(match.GoogleContact) != null));
                Assert.That(handler.Calls, Is.EqualTo(2));
                Assert.That(sync.ErrorCount, Is.EqualTo(recovers ? 0 : 1));
                Assert.That(errors, Is.EqualTo(sync.ErrorCount));
                Assert.That(sync.SyncedCount, Is.EqualTo(recovers ? 1 : 0));
            }
        }

        private static PeopleServiceService Service(FakeHandler handler) => new PeopleServiceService(
            new BaseClientService.Initializer { HttpClientFactory = new FakeFactory(handler) });

        private static Synchronizer SynchronizerFor(PeopleServiceService service)
        {
            var sync = new Synchronizer();
            typeof(Synchronizer).GetProperty(nameof(Synchronizer.GooglePeopleResource))
                .GetSetMethod(true).Invoke(sync, new object[] { service.People });
            return sync;
        }

        private sealed class FakeFactory : Google.Apis.Http.HttpClientFactory
        {
            private readonly HttpMessageHandler handler;
            internal FakeFactory(HttpMessageHandler handler) { this.handler = handler; }
            protected override HttpMessageHandler CreateHandler(CreateHttpClientArgs args) => handler;
        }

        private sealed class FakeHandler : HttpMessageHandler
        {
            private readonly Func<int, HttpResponseMessage> respond;
            internal int Calls { get; private set; }
            internal FakeHandler(Func<int, HttpResponseMessage> respond) { this.respond = respond; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                try { return Task.FromResult(respond(++Calls)); }
                catch (Exception ex) { return Task.FromException<HttpResponseMessage>(ex); }
            }
        }
    }
}
