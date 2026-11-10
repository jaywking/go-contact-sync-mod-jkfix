using Google.Apis.PeopleService.v1.Data;
using NUnit.Framework;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    [NonParallelizable]
    public class UnreadableContactTests
    {
        private sealed class Capture : ILogEventSink
        {
            internal readonly List<string> Warnings = new List<string>();
            public void Emit(LogEvent logEvent)
            {
                if (logEvent.Level == LogEventLevel.Warning) Warnings.Add(logEvent.RenderMessage());
            }
        }

        private delegate void FailedContact(Synchronizer sync, List<string> skippedIds, int index, int total);
        private ILogger previousLogger;
        private Logger logger;
        private Capture capture;
        private static COMException Failure() => new COMException("Original Outlook conflict", unchecked((int)0x80040119));

        [SetUp]
        public void SetUp()
        {
            previousLogger = Log.Logger;
            capture = new Capture();
            logger = new LoggerConfiguration().WriteTo.Sink(capture).CreateLogger();
            Log.Logger = logger;
        }

        [TearDown]
        public void TearDown()
        {
            Log.Logger = previousLogger;
            logger.Dispose();
        }

        private static string Read(string value)
        {
            if (value == null) throw Failure();
            return value;
        }

        private static FailedContact Contact(string id = "broken-id", string fileAs = "Jane Smith",
            string fullName = "Jane Full Name", string subject = "Jane Subject", bool folderFails = false,
            Exception emailFailure = null, Action onRead = null, Exception nameFailure = null)
        {
            return (sync, skippedIds, index, total) =>
            {
                onRead?.Invoke();
                ContactsMatcher.SkipUnreadableOutlookContact(sync, skippedIds, index, total, emailFailure ?? Failure(),
                    () => { if (nameFailure != null) throw nameFailure; return Read(fileAs); },
                    () => Read(fullName), () => Read(subject),
                    () => Read(folderFails ? null : "\\\\Mailbox\\Contacts"), () => Read(id));
            };
        }

        private static Person Google(string id, string outlookId = null) => new Person
        {
            ResourceName = "people/" + id,
            Names = new List<Name> { new Name
            {
                DisplayName = id, UnstructuredName = id,
                Metadata = new FieldMetadata { Primary = true, Source = new Source { Type = "CONTACT" } }
            } },
            Metadata = new PersonMetadata { Deleted = false },
            ClientData = outlookId == null ? null : new List<ClientData>
            {
                new ClientData { Key = OutlookPropertiesUtils.GetKey(), Value = outlookId }
            }
        };

        private static Synchronizer Sync(params Person[] google)
        {
            var sync = new Synchronizer { SyncDelete = true, SyncOption = SyncOption.MergeGoogleWins };
            typeof(Synchronizer).GetProperty(nameof(Synchronizer.GoogleContacts)).GetSetMethod(true)
                .Invoke(sync, new object[] { new Collection<Person>(google) });
            return sync;
        }

        private static List<ContactMatch> Match(Synchronizer sync, params FailedContact[] contacts)
        {
            // Keep production matching/deletion exclusion, replacing only COM property access.
            return ContactsMatcher.MatchContacts(sync, out _, (result, skippedIds, unmatched) =>
            {
                for (var i = 0; i < contacts.Length; i++) contacts[i](sync, skippedIds, i + 1, contacts.Length);
            });
        }

        [TestCase("Jane Smith", "Jane Full Name", "Jane Subject", "Jane Smith")]
        [TestCase(null, "Jane Full Name", "Jane Subject", "Jane Full Name")]
        [TestCase(" ", "Jane Full Name", "Jane Subject", "Jane Full Name")]
        [TestCase(null, null, "Jane Subject", "Jane Subject")]
        [TestCase(null, null, null, "(name unavailable)")]
        public void WarningRetainsBestIdentityLocationAndOriginalFailure(string fileAs, string fullName, string subject, string expected)
        {
            var sync = Sync(Google("peer", "broken-id"));
            var matches = Match(sync, Contact(fileAs: fileAs, fullName: fullName, subject: subject));
            Assert.That(matches, Is.Empty, "The failed Outlook contact and its linked Google peer must be excluded.");
            Assert.That(capture.Warnings[0], Does.Contain(expected).And.Contain("Mailbox").And.Contain("item 1 of 1")
                .And.Contain("Email1Address").And.Contain("0x80040119").And.Contain("Original Outlook conflict"));
            Assert.That(sync.SkippedCount, Is.EqualTo(1));
            Assert.That(sync.SkippedCountNotMatches, Is.EqualTo(1));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void MissingEntryIdPreservesNameAndStopsMatchingBeforeSaving(bool syncDelete)
        {
            var laterRead = false;
            var original = Failure();
            var sync = Sync(Google("peer", "broken-id"));
            sync.SyncDelete = syncDelete;
            var error = Assert.Throws<InvalidOperationException>(() => Match(sync,
                Contact(id: null, emailFailure: original), Contact(onRead: () => laterRead = true)));
            Assert.That(error.InnerException, Is.SameAs(original));
            Assert.That(error.Message, Does.Contain("Jane Smith").And.Contain("EntryID is unavailable").And.Contain("Original Outlook conflict"));
            Assert.That(laterRead, Is.False);
            Assert.That(capture.Warnings.Count, Is.EqualTo(1));
            Assert.That(sync.DeletedCount, Is.Zero);
            Assert.That(sync.SyncedCount, Is.Zero);
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase(" ")]
        public void MissingOrBlankEntryIdFailsClosedEvenWhenAllDiagnosticReadsFail(string id)
        {
            var sync = Sync();
            var error = Assert.Throws<InvalidOperationException>(() => Match(sync, Contact(id, null, null, null, true)));
            Assert.That(error.Message, Does.Contain("name unavailable").And.Contain("folder 'unavailable'").And.Contain("item 1 of 1"));
        }

        [Test]
        public void ReadableIdProtectsGooglePeerAndOtherContactsContinue()
        {
            var sync = Sync(Google("peer", "broken-id"), Google("second-peer", "second-id"), Google("healthy"));
            var matches = Match(sync, Contact(), Contact("second-id", "Second"));
            Assert.That(matches.Count, Is.EqualTo(1));
            Assert.That(matches[0].GoogleContact.ResourceName, Is.EqualTo("people/healthy"));
            var saved = new List<ContactMatch>();
            sync.SaveContacts(matches, match => { saved.Add(match); return true; });
            Assert.That(saved, Is.EqualTo(matches));
            Assert.That(sync.DeletedCount, Is.Zero);
            Assert.That(sync.SkippedCount, Is.EqualTo(2));
        }

        [Test]
        public void FolderFailureAndNewlinesDoNotHideIdentityOrForgeLogLines()
        {
            var sync = Sync();
            Match(sync, Contact(fileAs: "Jane\r\nSmith", folderFails: true,
                emailFailure: new COMException("Original\r\nconflict", unchecked((int)0x80040119))));
            Assert.That(capture.Warnings[0], Does.Contain("Jane  Smith").And.Contain("folder 'unavailable'").And.Contain("Original  conflict"));
            Assert.That(capture.Warnings[0], Does.Not.Contain("\r").And.Not.Contain("\n"));
        }

        [Test]
        public void CancellationDuringDiagnosticReadStopsMatching()
        {
            var sync = Sync();
            Assert.Throws<OperationCanceledException>(() => Match(sync, Contact(nameFailure: new OperationCanceledException())));
            Assert.That(sync.SkippedCount, Is.Zero);
            Assert.That(capture.Warnings, Is.Empty);
        }
    }
}
