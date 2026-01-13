using Google.Apis.Calendar.v3.Data;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using Serilog;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class SyncAppointmentsTests
    {
        private Synchronizer Synchronizer;

        //Constants for test appointment
        private readonly string name = OutlookAppointmentBuilder.testAppointmentSubject;

        [OneTimeSetUp]
        public void Init()
        {
            GoogleAPITests.LoadSettings(out var gmailUsername, out var syncProfile, out var syncContactsFolder, out var syncAppointmentsFolder);

            Synchronizer = new Synchronizer
            {
                SyncAppointments = true,
                SyncContacts = false
            };
            Synchronizer.SyncProfile = syncProfile;
            Assert.That(syncAppointmentsFolder, Is.Not.Null);;
            Synchronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
            Synchronizer.MonthsInPast = 1;
            Synchronizer.MonthsInFuture = 1;
            Synchronizer.RestrictMonthsInPast = true;
            Synchronizer.RestrictMonthsInFuture = true;

            Synchronizer.LoginToGoogle(gmailUsername);
            Synchronizer.LoginToOutlook();
        }

        [SetUp]
        public void SetUp()
        {
            DeleteTestAppointments();
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            Synchronizer.LogoffOutlook();
            Synchronizer.LogoffGoogle();
        }

        // Two synchronized non recurring appointments
        // On Google side appointment deleted
        // Testing if sync will create Google appointment
        [Test]
        public void Test_OutlookToGoogleOnly_With_DeletedGoogleAppointment()
        {
            var option = SyncOption.OutlookToGoogleOnly;

            Log.Information($"*** Test_OutlookToGoogleOnly_With_DeletedGoogleAppointment: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();
            oa.Start = new DateTime(2020, 9, 9, 12, 0, 0);
            oa.End = new DateTime(2020, 9, 9, 13, 0, 0);
            oa.Save();

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);
            ClassicAssert.IsNotNull(oid);

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            ClassicAssert.IsNotNull(gid);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            DeleteTestAppointment(ga);

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            Assert.That(!oa.IsRecurring);
            ClassicAssert.AreEqual(new DateTime(2020, 9, 9, 12, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 9, 9, 13, 0, 0), oa.End);

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(ga);
            ClassicAssert.AreEqual(new DateTime(2020, 9, 9, 12, 0, 0), ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 9, 9, 13, 0, 0), ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.IsNull(ga.Recurrence);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_OutlookToGoogleOnly_With_DeletedGoogleAppointment: {option} ***");
        }


        // Two synchronized non recurring appointments
        // On Outlook side appointment with multiple recipients
        // updated by changing end date of the appointment
        // Testing if sync will revert Outlook changes
        [Test]
        public void Test_Multiple_Recipients_Google_Wins()
        {
            var option = SyncOption.GoogleToOutlookOnly;

            Log.Information($"*** Test_Multiple_Recipients_Google_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();
            oa.Start = new DateTime(2020, 6, 20, 12, 0, 0);
            oa.End = new DateTime(2020, 6, 20, 13, 0, 0);
            var r = oa.Recipients;
            var r1 = r.Add("test01@example.com");
            r1.Type = (int)Outlook.OlMailRecipientType.olTo;
            var r2 = r.Add("test02@example.com");
            r2.Type = (int)Outlook.OlMailRecipientType.olTo;

            oa.Save();
            ClassicAssert.AreEqual(true, oa.Saved);
            ClassicAssert.AreEqual(2, r.Count);
            var dt1 = oa.LastModificationTime;
            if (r1 != null)
            {
                Marshal.ReleaseComObject(r1);
            }
            if (r2 != null)
            {
                Marshal.ReleaseComObject(r2);
            }
            if (r != null)
            {
                Marshal.ReleaseComObject(r);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);
            ClassicAssert.IsNotNull(oid);

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            ClassicAssert.IsNotNull(gid);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            oa.End = new DateTime(2020, 6, 20, 14, 0, 0);
            oa.Save();
            ClassicAssert.AreEqual(new DateTime(2020, 6, 20, 14, 0, 0), oa.End);
            ClassicAssert.AreEqual(true, oa.Saved);
            r = oa.Recipients;
            ClassicAssert.AreEqual(2, r.Count);
            var dt2 = oa.LastModificationTime;
            ClassicAssert.Greater(dt2, dt1);
            if (r != null)
            {
                Marshal.ReleaseComObject(r);
            }
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            Assert.That(!oa.IsRecurring);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 20, 12, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 20, 14, 0, 0), oa.End);

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            Synchronizer.LoadAppointments();
            //ClassicAssert.AreEqual(0, sync.appointmentsSynchronizer.GoogleAppointments.Count);
            //ClassicAssert.AreEqual(0, sync.appointmentsSynchronizer.OutlookAppointments.Count);
            ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(ga);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 20, 12, 0, 0), ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 20, 13, 0, 0), ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.IsNull(ga.Recurrence);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Multiple_Recipients_Google_Wins: {option} ***");

        }

        // Two synchronized non recurring appointments
        // On Outlook side appointment updated by changing 
        // start date of the appointment
        // Testing if sync will update Google appointment
        // Duration of tested appointment is for one hour
        [Test]
        public void Test_End_Date_Change_Outlook_Side_One_Hour_Outlook_Wins()
        {
            var option = SyncOption.OutlookToGoogleOnly;

            Log.Information($"*** Test_End_Date_Change_Outlook_Side_One_Hour_Outlook_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();
            oa.Start = new DateTime(2020, 6, 16, 12, 0, 0);
            oa.End = new DateTime(2020, 6, 16, 13, 0, 0);
            oa.Save();

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);
            ClassicAssert.IsNotNull(oid);

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            ClassicAssert.IsNotNull(gid);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            oa.End = new DateTime(2020, 6, 16, 14, 0, 0);
            oa.Save();
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            Assert.That(!oa.IsRecurring);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 16, 12, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 16, 14, 0, 0), oa.End);

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(ga);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 16, 12, 0, 0), ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 16, 14, 0, 0), ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.IsNull(ga.Recurrence);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_End_Date_Change_Outlook_Side_One_Hour_Outlook_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On Outlook side recurrence exception created by changing 
        // start date of the appointment
        // Testing if sync will update Google exception
        // Duration of tested appointment is for one hour
        [Test]
        public void Test_Start_Date_Change_Exceptions_Outlook_Side_One_Hour_Outlook_Wins()
        {
            var option = SyncOption.OutlookToGoogleOnly;

            Log.Information($"*** Test_Start_Date_Change_Exceptions_Outlook_Side_One_Hour_Outlook_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 6, 3, 12, 0, 0);
            rp.PatternEndDate = new DateTime(2020, 6, 10, 12, 0, 0);
            rp.Duration = 90;
            rp.StartTime = new DateTime(1899, 12, 30, 15, 0, 0);
            rp.EndTime = new DateTime(1899, 12, 30, 16, 30, 0);
            oa.Save();

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            ClassicAssert.IsNotNull(gid);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            var ex = rp.GetOccurrence(new DateTime(2020, 6, 3, 15, 0, 0));
            ClassicAssert.IsNotNull(ex);
            ex.Start = new DateTime(2020, 6, 3, 16, 0, 0);
            ex.Save();
            oa.Save();

            if (ex != null)
            {
                Marshal.ReleaseComObject(ex);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);
            ClassicAssert.IsNotNull(oid);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 15, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 30, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 0, 0, 0), rp.PatternStartDate);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 10, 0, 0, 0), rp.PatternEndDate);
            ClassicAssert.AreEqual(90, rp.Duration);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 15, 0, 0), rp.StartTime);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 16, 30, 0), rp.EndTime);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(1, exceptions.Count);

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 15, 0, 0), f_ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 30, 0), f_ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(1, f_ga.Recurrence.Count);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;UNTIL=20200611;BYDAY=WE", f_ga.Recurrence[0]);

            var instances = Synchronizer.GetGoogleAppointmentInstances(gid);
            instances.OriginalStart = "2020-06-03T15:00:00+02:00";
            var r = instances.Execute();
            ClassicAssert.AreEqual(1, r.Items.Count);
            var i1 = r.Items[0];

            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 0, 0), i1.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 17, 30, 0), i1.End.DateTimeDateTimeOffset.Value.DateTime);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Start_Date_Change_Exceptions_Outlook_Side_One_Hour_Outlook_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On Outlook side recurrence exception created by changing 
        // start date of the appointment
        // On Google side recurrence exception created by changing 
        // start date of the appointment
        // Testing if sync will update Google exception
        // Duration of tested appointment is for one hour
        [Test]
        public void Test_Start_Date_Change_Exceptions_Both_Side_One_Hour_Outlook_Wins()
        {
            var option = SyncOption.OutlookToGoogleOnly;

            Log.Information($"*** Test_Start_Date_Change_Exceptions_Both_Side_One_Hour_Outlook_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 6, 3, 12, 0, 0);
            rp.PatternEndDate = new DateTime(2020, 6, 10, 12, 0, 0);
            rp.Duration = 90;
            rp.StartTime = new DateTime(1899, 12, 30, 15, 0, 0);
            rp.EndTime = new DateTime(1899, 12, 30, 16, 30, 0);
            oa.Save();

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            ClassicAssert.IsNotNull(gid);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            var ex = rp.GetOccurrence(new DateTime(2020, 6, 3, 15, 0, 0));
            ClassicAssert.IsNotNull(ex);
            ex.Start = new DateTime(2020, 6, 3, 16, 0, 0);
            ex.Save();
            oa.Save();

            if (ex != null)
            {
                Marshal.ReleaseComObject(ex);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);
            ClassicAssert.IsNotNull(oid);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var instances = Synchronizer.GetGoogleAppointmentInstances(gid);
            instances.OriginalStart = "2020-06-03T15:00:00+02:00";
            var r = instances.Execute();
            ClassicAssert.AreEqual(1, r.Items.Count);
            var i1 = r.Items[0];

            i1.Start = new EventDateTime
            {
                DateTimeDateTimeOffset = new DateTimeOffset(new DateTime(2020, 6, 3, 19, 0, 0)),
                TimeZone = "Europe/Warsaw"
            };
            i1.End = new EventDateTime
            {
                DateTimeDateTimeOffset = new DateTimeOffset(new DateTime(2020, 6, 3, 20, 0, 0)),
                TimeZone = "Europe/Warsaw"
            };

            var d_ga = Synchronizer.SaveGoogleAppointment(i1);
            var d_gid = AppointmentPropertiesUtils.GetGoogleId(d_ga);
            ClassicAssert.IsNotNull(d_gid);

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 15, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 30, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 0, 0, 0), rp.PatternStartDate);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 10, 0, 0, 0), rp.PatternEndDate);
            ClassicAssert.AreEqual(90, rp.Duration);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 15, 0, 0), rp.StartTime);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 16, 30, 0), rp.EndTime);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(1, exceptions.Count);

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 15, 0, 0), f_ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 30, 0), f_ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(1, f_ga.Recurrence.Count);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;UNTIL=20200611;BYDAY=WE", f_ga.Recurrence[0]);

            instances = Synchronizer.GetGoogleAppointmentInstances(gid);
            instances.OriginalStart = "2020-06-03T15:00:00+02:00";
            r = instances.Execute();
            ClassicAssert.AreEqual(1, r.Items.Count);
            i1 = r.Items[0];

            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 0, 0), i1.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 17, 30, 0), i1.End.DateTimeDateTimeOffset.Value.DateTime);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Start_Date_Change_Exceptions_Both_Side_One_Hour_Outlook_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On Outlook side recurrence exception created by changing 
        // start date of the appointment
        // Testing if sync will update Outlook exception
        // Duration of tested appointment is for one hour
        [Test]
        public void Test_Start_Date_Change_Exceptions_Outlook_Side_One_Hour_Google_Wins()
        {
            var option = SyncOption.GoogleToOutlookOnly;

            Log.Information($"*** Test_Start_Date_Change_Exceptions_Outlook_Side_One_Hour_Google_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 6, 3, 12, 0, 0);
            rp.PatternEndDate = new DateTime(2020, 6, 10, 12, 0, 0);
            rp.Duration = 90;
            rp.StartTime = new DateTime(1899, 12, 30, 15, 0, 0);
            rp.EndTime = new DateTime(1899, 12, 30, 16, 30, 0);
            oa.Save();

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            ClassicAssert.IsNotNull(gid);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            var ex = rp.GetOccurrence(new DateTime(2020, 6, 3, 15, 0, 0));
            ClassicAssert.IsNotNull(ex);
            ex.Start = new DateTime(2020, 6, 3, 16, 0, 0);
            ex.Save();
            oa.Save();

            if (ex != null)
            {
                Marshal.ReleaseComObject(ex);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);
            ClassicAssert.IsNotNull(oid);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 15, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 30, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 0, 0, 0), rp.PatternStartDate);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 10, 0, 0, 0), rp.PatternEndDate);
            ClassicAssert.AreEqual(90, rp.Duration);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 15, 0, 0), rp.StartTime);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 16, 30, 0), rp.EndTime);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(0, exceptions.Count);

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 15, 0, 0), f_ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 30, 0), f_ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(1, f_ga.Recurrence.Count);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;UNTIL=20200611;BYDAY=WE", f_ga.Recurrence[0]);

            var insn = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.AreEqual(2, insn.Items.Count);
            var j1 = insn.Items[0];
            ClassicAssert.IsNotNull(j1);
            var j2 = insn.Items[1];
            ClassicAssert.IsNotNull(j2);
            //instances could be returned in not sorted order
            if (j2.Start.DateTimeDateTimeOffset.Value < j1.Start.DateTimeDateTimeOffset.Value.DateTime)
            {
                var t = j2;
                j2 = j1;
                j1 = t;
            }
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 15, 0, 0), j1.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 3, 16, 30, 0), j1.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 10, 15, 0, 0), j2.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 6, 10, 16, 30, 0), j2.End.DateTimeDateTimeOffset.Value.DateTime);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Start_Date_Change_Exceptions_Outlook_Side_One_Hour_Google_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On Outlook side recurrence exception created by deleting 
        // one instance of the appointment
        // Testing if sync will undelete deleted Outlook exception
        // Duration of tested appointment is for one hour
        [Test]
        public void Test_Deleted_Exceptions_Outlook_Side_One_Hour_Google_Wins()
        {
            var option = SyncOption.GoogleToOutlookOnly;

            Log.Information($"*** Test_Deleted_Exceptions_Outlook_Side_One_Hour_Google_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 1, 8, 12, 0, 0);
            rp.PatternEndDate = new DateTime(2020, 1, 15, 12, 0, 0);
            rp.Duration = 90;
            rp.StartTime = new DateTime(1899, 12, 30, 15, 0, 0);
            rp.EndTime = new DateTime(1899, 12, 30, 16, 30, 0);

            oa.Save();
            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            rp = oa.GetRecurrencePattern();
            var ex1 = rp.GetOccurrence(new DateTime(2020, 1, 8, 15, 0, 0));
            ex1.Delete();

            oa.Save();

            if (ex1 != null)
            {
                Marshal.ReleaseComObject(ex1);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), rp.PatternStartDate);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 0, 0, 0), rp.PatternEndDate);
            ClassicAssert.AreEqual(90, rp.Duration);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 15, 0, 0), rp.StartTime);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 16, 30, 0), rp.EndTime);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(0, exceptions.Count);

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), f_ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), f_ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.Greater(f_ga.Recurrence.Count, 0);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;UNTIL=20200116;BYDAY=WE", f_ga.Recurrence[0]);

            var insn = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.AreEqual(2, insn.Items.Count);
            var j1 = insn.Items[0];
            ClassicAssert.IsNotNull(j1);
            var j2 = insn.Items[1];
            ClassicAssert.IsNotNull(j2);
            //instances could be returned in not sorted order
            if (j2.Start.DateTimeDateTimeOffset.Value < j1.Start.DateTimeDateTimeOffset.Value.DateTime)
            {
                var t = j2;
                j2 = j1;
                j1 = t;
            }
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), j1.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), j1.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 15, 0, 0), j2.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 16, 30, 0), j2.End.DateTimeDateTimeOffset.Value.DateTime);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Deleted_Exceptions_Outlook_Side_One_Hour_Google_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On Outlook side recurrence exception created by deleting 
        // one instance of the appointment
        // Testing if sync will not change anything
        // Tested appointment is for one hour
        [Test]
        public void Test_Deleted_Exceptions_Outlook_Side_One_Hour_Outlook_Wins()
        {
            var option = SyncOption.OutlookToGoogleOnly;

            Log.Information($"*** Test_Deleted_Exceptions_Outlook_Side_One_Hour_Outlook_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 5, 6, 12, 0, 0);
            rp.NoEndDate = true;
            rp.Duration = 90;
            rp.StartTime = new DateTime(1899, 12, 30, 15, 0, 0);
            rp.EndTime = new DateTime(1899, 12, 30, 16, 30, 0);

            oa.Save();
            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            rp = oa.GetRecurrencePattern();
            var ex1 = rp.GetOccurrence(new DateTime(2020, 5, 6, 15, 0, 0));
            ex1.Delete();

            oa.Save();

            if (ex1 != null)
            {
                Marshal.ReleaseComObject(ex1);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 6, 15, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 6, 16, 30, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 6, 0, 0, 0), rp.PatternStartDate);
            Assert.That(rp.NoEndDate);
            ClassicAssert.AreEqual(90, rp.Duration);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 15, 0, 0), rp.StartTime);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 16, 30, 0), rp.EndTime);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(1, exceptions.Count);

            var e = exceptions[1];
            ClassicAssert.IsNotNull(e);
            Assert.That(e.Deleted);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 6, 0, 0, 0), e.OriginalDate);

            if (e != null)
            {
                Marshal.ReleaseComObject(e);
            }

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 6, 15, 0, 0), f_ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 6, 16, 30, 0), f_ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.Greater(f_ga.Recurrence.Count, 0);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;BYDAY=WE", f_ga.Recurrence[0]);

            var insn = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.Greater(insn.Items.Count, 0);
            var i = insn.Items[0];
            ClassicAssert.IsNotNull(i);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 13, 15, 0, 0), i.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 13, 16, 30, 0), i.End.DateTimeDateTimeOffset.Value.DateTime);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Deleted_Exceptions_Outlook_Side_One_Hour_Outlook_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On Outlook side recurrence exception created by deleting 
        // one instance of the appointment
        // Testing if sync will not change anything
        // Tested appointment is for full day
        [Test]
        public void Test_Deleted_Exceptions_Outlook_Side_Full_Day_Outlook_Wins()
        {
            var option = SyncOption.OutlookToGoogleOnly;

            Log.Information($"*** Test_Deleted_Exceptions_Outlook_Side_Full_Day_Outlook_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildAllDay();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 5, 6, 0, 0, 0);
            rp.NoEndDate = true;

            oa.Save();
            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            rp = oa.GetRecurrencePattern();
            var ex1 = rp.GetOccurrence(new DateTime(2020, 5, 20, 0, 0, 0));
            ex1.Delete();

            oa.Save();

            if (ex1 != null)
            {
                Marshal.ReleaseComObject(ex1);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 6, 0, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 7, 0, 0, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 6, 0, 0, 0), rp.PatternStartDate);
            Assert.That(rp.NoEndDate);
            ClassicAssert.AreEqual(1440, rp.Duration);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(1, exceptions.Count);

            var e = exceptions[1];
            ClassicAssert.IsNotNull(e);
            Assert.That(e.Deleted);
            ClassicAssert.AreEqual(new DateTime(2020, 5, 20, 0, 0, 0), e.OriginalDate);

            if (e != null)
            {
                Marshal.ReleaseComObject(e);
            }

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual("2020-05-06", f_ga.Start.Date);
            ClassicAssert.AreEqual("2020-05-07", f_ga.End.Date);
            ClassicAssert.Greater(f_ga.Recurrence.Count, 0);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;BYDAY=WE", f_ga.Recurrence[0]);

            var r = Synchronizer.GetGoogleAppointmentInstances(gid);
            r.OriginalStart = "2020-05-20";
            r.ShowDeleted = true;
            var insn = r.Execute();

            ClassicAssert.AreEqual(1, insn.Items.Count);
            var i = insn.Items[0];
            ClassicAssert.IsNotNull(i);
            ClassicAssert.AreEqual("2020-05-20", i.Start.Date);
            ClassicAssert.AreEqual("2020-05-21", i.End.Date);
            ClassicAssert.AreEqual("cancelled", i.Status);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Deleted_Exceptions_Outlook_Side_Full_Day_Outlook_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On Google side recurrence exception created by deleting 
        // one instance of the appointment
        // Testing if sync will not change anything
        // Tested appointment is for one hour
        [Test]
        public void Test_Deleted_Exceptions_Google_Side_One_Hour_Outlook_Wins()
        {
            var option = SyncOption.OutlookToGoogleOnly;

            Log.Information($"*** Test_Deleted_Exceptions_Google_Side_One_Hour_Outlook_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 1, 8, 12, 0, 0);
            rp.PatternEndDate = new DateTime(2020, 1, 15, 12, 0, 0);
            rp.Duration = 90;
            rp.StartTime = new DateTime(1899, 12, 30, 15, 0, 0);
            rp.EndTime = new DateTime(1899, 12, 30, 16, 30, 0);

            oa.Save();
            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var instances = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.AreEqual(2, instances.Items.Count);
            var i1 = instances.Items[0];
            ClassicAssert.IsNotNull(i1);
            ClassicAssert.IsNotNull(i1.Start);
            ClassicAssert.IsNotNull(i1.Start.DateTimeDateTimeOffset.Value);
            var i2 = instances.Items[1];
            ClassicAssert.IsNotNull(i2);
            ClassicAssert.IsNotNull(i2.Start);
            ClassicAssert.IsNotNull(i2.Start.DateTimeDateTimeOffset.Value);

            //instances could be returned in not sorted order
            if (i2.Start.DateTimeDateTimeOffset.Value < i1.Start.DateTimeDateTimeOffset.Value)
            {
                var t = i2;
                i2 = i1;
                i1 = t;
            }

            var i1_gid = AppointmentPropertiesUtils.GetGoogleId(i1);
            ClassicAssert.IsNotNull(i1_gid);
            var i2_gid = AppointmentPropertiesUtils.GetGoogleId(i2);
            ClassicAssert.IsNotNull(i2_gid);

            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), i1.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 15, 0, 0), i2.Start.DateTimeDateTimeOffset.Value.DateTime);

            i1.Status = "cancelled";
            var d_ga = Synchronizer.SaveGoogleAppointment(i1);
            var d_gid = AppointmentPropertiesUtils.GetGoogleId(d_ga);
            ClassicAssert.IsNotNull(d_gid);

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), rp.PatternStartDate);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 0, 0, 0), rp.PatternEndDate);
            ClassicAssert.AreEqual(90, rp.Duration);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 15, 0, 0), rp.StartTime);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 16, 30, 0), rp.EndTime);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(0, exceptions.Count);

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), f_ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), f_ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.Greater(f_ga.Recurrence.Count, 0);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;UNTIL=20200116;BYDAY=WE", f_ga.Recurrence[0]);

            var ir = Synchronizer.GetGoogleAppointmentInstances(gid);
            ir.ShowDeleted = true;
            var insn = ir.Execute();
            ClassicAssert.AreEqual(2, insn.Items.Count);
            var j1 = insn.Items[0];
            ClassicAssert.IsNotNull(j1);
            var j2 = insn.Items[1];
            ClassicAssert.IsNotNull(j2);
            //instances could be returned in not sorted order
            if (j2.Start.DateTimeDateTimeOffset.Value < j1.Start.DateTimeDateTimeOffset.Value)
            {
                var t = j2;
                j2 = j1;
                j1 = t;
            }
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), j1.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), j1.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 15, 0, 0), j2.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 16, 30, 0), j2.End.DateTimeDateTimeOffset.Value.DateTime);

            var f_d_ga = Synchronizer.GetGoogleAppointment(d_gid);
            ClassicAssert.IsNotNull(f_d_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), f_d_ga.OriginalStartTime.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual("confirmed", f_d_ga.Status);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Deleted_Exceptions_Google_Side_One_Hour_Outlook_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On Google side recurrence exception created by deleting 
        // one instance of the appointment
        // Testing if sync will not change anything
        // Tested appointment is for one hour
        [Test]
        public void Test_Deleted_Exceptions_Google_Side_One_Hour_Google_Wins()
        {
            var option = SyncOption.GoogleToOutlookOnly;

            Log.Information($"*** Test_Deleted_Exceptions_Google_Side_One_Hour_Google_Wins: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 1, 8, 12, 0, 0);
            rp.NoEndDate = true;
            rp.Duration = 90;
            rp.StartTime = new DateTime(1899, 12, 30, 15, 0, 0);
            rp.EndTime = new DateTime(1899, 12, 30, 16, 30, 0);

            oa.Save();
            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var instances = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.Greater(instances.Items.Count, 0);

            var instance = instances.Items[0];
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), instance.Start.DateTimeDateTimeOffset.Value.DateTime);

            instance.Status = "cancelled";
            var d_ga = Synchronizer.SaveGoogleAppointment(instance);
            var d_gid = AppointmentPropertiesUtils.GetGoogleId(d_ga);
            ClassicAssert.IsNotNull(d_gid);

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), rp.PatternStartDate);
            Assert.That(rp.NoEndDate);
            ClassicAssert.AreEqual(90, rp.Duration);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 15, 0, 0), rp.StartTime);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 16, 30, 0), rp.EndTime);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(1, exceptions.Count);

            var e = exceptions[1];
            ClassicAssert.IsNotNull(e);
            Assert.That(e.Deleted);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), e.OriginalDate);

            if (e != null)
            {
                Marshal.ReleaseComObject(e);
            }

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), f_ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), f_ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.Greater(f_ga.Recurrence.Count, 0);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;BYDAY=WE", f_ga.Recurrence[0]);

            var insn = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.Greater(insn.Items.Count, 0);
            var i = insn.Items[0];
            ClassicAssert.IsNotNull(i);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 15, 0, 0), i.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 16, 30, 0), i.End.DateTimeDateTimeOffset.Value.DateTime);

            var f_d_ga = Synchronizer.GetGoogleAppointmentById(d_gid);
            ClassicAssert.IsNotNull(f_d_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), f_d_ga.OriginalStartTime.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual("cancelled", f_d_ga.Status);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Deleted_Exceptions_Google_Side_One_Hour_Google_Wins: {option} ***");
        }

        // Two recurring synchronized appointments
        // On both sides (Google and Outlook) recurrence exception created by deleting 
        // one instance of the appointment
        // Testing if sync will not change anything
        // Tested appointment is for one hour
        [Test]
        public void Test_Deleted_Exceptions_Both_Side_One_Hour([Values(SyncOption.GoogleToOutlookOnly, SyncOption.OutlookToGoogleOnly)] SyncOption option)
        {
            Log.Information($"*** Test_Deleted_Exceptions_Both_Side_One_Hour: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildDefault();

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 1, 8, 12, 0, 0);
            rp.NoEndDate = true;
            rp.Duration = 90;
            rp.StartTime = new DateTime(1899, 12, 30, 15, 0, 0);
            rp.EndTime = new DateTime(1899, 12, 30, 16, 30, 0);

            oa.Save();
            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            rp = oa.GetRecurrencePattern();
            var ex1 = rp.GetOccurrence(new DateTime(2020, 1, 8, 15, 0, 0));
            ex1.Delete();

            oa.Save();

            if (ex1 != null)
            {
                Marshal.ReleaseComObject(ex1);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var instances = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.Greater(instances.Items.Count, 0);

            var instance = instances.Items[0];
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), instance.Start.DateTimeDateTimeOffset.Value.DateTime);

            instance.Status = "cancelled";
            var d_ga = Synchronizer.SaveGoogleAppointment(instance);
            var d_gid = AppointmentPropertiesUtils.GetGoogleId(d_ga);
            ClassicAssert.IsNotNull(d_gid);

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), rp.PatternStartDate);
            Assert.That(rp.NoEndDate);
            ClassicAssert.AreEqual(90, rp.Duration);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 15, 0, 0), rp.StartTime);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 16, 30, 0), rp.EndTime);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(1, exceptions.Count);

            var e = exceptions[1];
            ClassicAssert.IsNotNull(e);
            Assert.That(e.Deleted);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), e.OriginalDate);

            if (e != null)
            {
                Marshal.ReleaseComObject(e);
            }

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), f_ga.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 16, 30, 0), f_ga.End.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.Greater(f_ga.Recurrence.Count, 0);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;BYDAY=WE", f_ga.Recurrence[0]);

            var insn = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.Greater(insn.Items.Count, 0);
            var i = insn.Items[0];
            ClassicAssert.IsNotNull(i);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 15, 0, 0), i.Start.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 16, 30, 0), i.End.DateTimeDateTimeOffset.Value.DateTime);

            var f_d_ga = Synchronizer.GetGoogleAppointmentById(d_gid);
            ClassicAssert.IsNotNull(f_d_ga);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 15, 0, 0), f_d_ga.OriginalStartTime.DateTimeDateTimeOffset.Value.DateTime);
            ClassicAssert.AreEqual("cancelled", f_d_ga.Status);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Deleted_Exceptions_Both_Side_One_Hour: {option} ***");
        }

        // Two recurring synchronized appointments
        // On both sides (Google and Outlook) recurrence exception created by deleting 
        // one instance of the appointment
        // Testing if sync will not change anything
        // Tested appointment is for full day
        [Test]
        public void Test_Deleted_Exceptions_Both_Side_Full_Day([Values(SyncOption.GoogleToOutlookOnly, SyncOption.OutlookToGoogleOnly)] SyncOption option)
        {
            Log.Information($"*** Test_Deleted_Exceptions_Both_Side_Full_Day: {option} ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa = new OutlookAppointmentBuilder().BuildAllDay();
            oa.AllDayEvent = true;

            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;
            rp.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday;
            rp.PatternStartDate = new DateTime(2020, 1, 8, 0, 0, 0);
            rp.NoEndDate = true;

            oa.Save();
            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            var gid = AppointmentPropertiesUtils.GetGoogleId(ga);
            Assert.That(AppointmentSync.IsSameRecurrence(ga, oa));

            DelayBetweenSync();

            rp = oa.GetRecurrencePattern();
            var ex1 = rp.GetOccurrence(new DateTime(2020, 1, 8, 0, 0, 0));
            ex1.Delete();

            oa.Save();

            if (ex1 != null)
            {
                Marshal.ReleaseComObject(ex1);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var instances = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.Greater(instances.Items.Count, 0);

            var instance = instances.Items[0];
            ClassicAssert.IsNotNull(instance.Start.Date);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), DateTime.Parse(instance.Start.Date));

            instance.Status = "cancelled";
            var d_ga = Synchronizer.SaveGoogleAppointment(instance);
            var d_gid = AppointmentPropertiesUtils.GetGoogleId(d_ga);
            ClassicAssert.IsNotNull(d_gid);

            DelayBetweenSync();

            //Act
            Synchronizer.SyncOption = option;
            Synchronizer.RestrictMonthsInPast = false;
            Synchronizer.RestrictMonthsInFuture = false;
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            CleanAllLoadedAppointments();

            //Assert
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), oa.Start);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 9, 0, 0, 0), oa.End);
            Assert.That(oa.IsRecurring);

            rp = oa.GetRecurrencePattern();
            ClassicAssert.IsNotNull(rp);
            ClassicAssert.AreEqual(Outlook.OlRecurrenceType.olRecursWeekly, rp.RecurrenceType);
            ClassicAssert.AreEqual(Outlook.OlDaysOfWeek.olWednesday, rp.DayOfWeekMask);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), rp.PatternStartDate);
            Assert.That(rp.NoEndDate);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 0, 0, 0).TimeOfDay, rp.StartTime.TimeOfDay);
            ClassicAssert.AreEqual(new DateTime(1899, 12, 30, 0, 0, 0).TimeOfDay, rp.EndTime.TimeOfDay);

            var exceptions = rp.Exceptions;
            ClassicAssert.IsNotNull(exceptions);
            ClassicAssert.AreEqual(1, exceptions.Count);

            var e = exceptions[1];
            ClassicAssert.IsNotNull(e);
            Assert.That(e.Deleted);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), e.OriginalDate);

            if (e != null)
            {
                Marshal.ReleaseComObject(e);
            }

            if (exceptions != null)
            {
                Marshal.ReleaseComObject(exceptions);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            DeleteTestAppointment(oa);
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            var f_ga = Synchronizer.GetGoogleAppointment(gid);
            ClassicAssert.IsNotNull(f_ga);
            ClassicAssert.IsNotNull(f_ga.Start.Date);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), DateTime.Parse(f_ga.Start.Date));
            ClassicAssert.IsNotNull(f_ga.End.Date);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 9, 0, 0, 0), DateTime.Parse(f_ga.End.Date));
            ClassicAssert.Greater(f_ga.Recurrence.Count, 0);
            ClassicAssert.AreEqual("RRULE:FREQ=WEEKLY;BYDAY=WE", f_ga.Recurrence[0]);

            var insn = Synchronizer.GetGoogleAppointmentInstances(gid).Execute();
            ClassicAssert.Greater(insn.Items.Count, 0);
            var i = insn.Items[0];
            ClassicAssert.IsNotNull(i);
            ClassicAssert.IsNotNull(i.Start.Date);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 15, 0, 0, 0), DateTime.Parse(i.Start.Date));
            ClassicAssert.IsNotNull(i.End.Date);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 16, 0, 0, 0), DateTime.Parse(i.End.Date));

            var f_d_ga = Synchronizer.GetGoogleAppointmentById(d_gid);
            ClassicAssert.IsNotNull(f_d_ga);
            ClassicAssert.IsNotNull(f_d_ga.OriginalStartTime.Date);
            ClassicAssert.AreEqual(new DateTime(2020, 1, 8, 0, 0, 0), DateTime.Parse(f_d_ga.OriginalStartTime.Date));
            ClassicAssert.AreEqual("cancelled", f_d_ga.Status);

            DeleteTestAppointment(ga);

            Log.Information($"*** Test_Deleted_Exceptions_Both_Side_Full_Day: {option} ***");
        }

        [Test]
        public void TestRemoveGoogleDuplicatedAppointments_01()
        {
            Log.Information($"*** TestRemoveGoogleDuplicatedAppointments_01 ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa1 = new OutlookAppointmentBuilder().BuildDefault();
            oa1.Save();

            // create new Google test appointments
            var ga1 = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa1, ref ga1);
            var ga2 = new GoogleAppointmentBuilder().Build();
            AppointmentSync.UpdateAppointment(oa1, ga2);
            AppointmentPropertiesUtils.SetGoogleOutlookId(ga2, oa1);
            ga2 = Synchronizer.SaveGoogleAppointment(ga2);

            var gid_oa1 = AppointmentPropertiesUtils.GetOutlookGoogleId(oa1);
            var gid_ga1 = AppointmentPropertiesUtils.GetGoogleId(ga1);
            var gid_ga2 = AppointmentPropertiesUtils.GetGoogleId(ga2);
            var oid_ga1 = AppointmentPropertiesUtils.GetGoogleOutlookId(ga1);
            var oid_ga2 = AppointmentPropertiesUtils.GetGoogleOutlookId(ga2);
            var oid_oa1 = AppointmentPropertiesUtils.GetOutlookId(oa1);

            if (oa1 != null)
            {
                Marshal.ReleaseComObject(oa1);
            }

            // assert appointments oa1 and ga1 are pointing to each other
            ClassicAssert.AreEqual(gid_oa1, gid_ga1);
            ClassicAssert.AreEqual(oid_oa1, oid_ga1);
            // assert appointment ga2 also points to oa1
            ClassicAssert.AreEqual(oid_oa1, oid_ga2);
            // assert appointment oa1 does not point to ga2
            ClassicAssert.AreNotEqual(gid_oa1, gid_ga2);

            //Act
            Synchronizer.LoadAppointments();

            //Assert
            var f_ga1 = Synchronizer.GetGoogleAppointmentById(gid_ga1);
            var f_ga2 = Synchronizer.GetGoogleAppointmentById(gid_ga2);
            var f_oa1 = Synchronizer.GetOutlookAppointmentById(oid_oa1);

            ClassicAssert.IsNotNull(f_ga1);
            ClassicAssert.IsNotNull(f_ga2);
            ClassicAssert.IsNotNull(f_oa1);

            var f_gid_oa1 = AppointmentPropertiesUtils.GetOutlookGoogleId(f_oa1);
            var f_gid_ga1 = AppointmentPropertiesUtils.GetGoogleId(f_ga1);
            var f_gid_ga2 = AppointmentPropertiesUtils.GetGoogleId(f_ga2);
            var f_oid_ga1 = AppointmentPropertiesUtils.GetGoogleOutlookId(f_ga1);
            var f_oid_ga2 = AppointmentPropertiesUtils.GetGoogleOutlookId(f_ga2);
            var f_oid_oa1 = AppointmentPropertiesUtils.GetOutlookId(f_oa1);

            // assert appointments oa and ga1 are pointing to each other
            ClassicAssert.AreEqual(f_gid_oa1, f_gid_ga1);
            ClassicAssert.AreEqual(f_oid_oa1, f_oid_ga1);
            // assert appointment ga2 does not point to oa
            ClassicAssert.AreNotEqual(f_oid_oa1, f_oid_ga2);
            // assert appointment oa1 does not point to ga2
            ClassicAssert.AreNotEqual(f_gid_oa1, f_gid_ga2);

            DeleteTestAppointment(f_oa1);
            DeleteTestAppointment(f_ga1);
            DeleteTestAppointment(f_ga2);

            if (f_oa1 != null)
            {
                Marshal.ReleaseComObject(f_oa1);
            }

            Log.Information($"*** TestRemoveGoogleDuplicatedAppointments_01 ***");
        }

        [Test]
        public void TestRemoveGoogleDuplicatedAppointments_02()
        {
            Log.Information($"*** TestRemoveGoogleDuplicatedAppointments_02 ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa1 = new OutlookAppointmentBuilder().BuildDefault();
            oa1.Save();

            // create new Google test appointments
            var ga1 = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa1, ref ga1);
            var ga2 = new GoogleAppointmentBuilder().Build();
            AppointmentSync.UpdateAppointment(oa1, ga2);
            AppointmentPropertiesUtils.SetGoogleOutlookId(ga2, oa1);
            ga2 = Synchronizer.SaveGoogleAppointment(ga2);
            AppointmentPropertiesUtils.ResetOutlookGoogleId(oa1);
            oa1.Save();

            var gid_oa1 = AppointmentPropertiesUtils.GetOutlookGoogleId(oa1);
            var gid_ga1 = AppointmentPropertiesUtils.GetGoogleId(ga1);
            var gid_ga2 = AppointmentPropertiesUtils.GetGoogleId(ga2);
            var oid_ga1 = AppointmentPropertiesUtils.GetGoogleOutlookId(ga1);
            var oid_ga2 = AppointmentPropertiesUtils.GetGoogleOutlookId(ga2);
            var oid_oa1 = AppointmentPropertiesUtils.GetOutlookId(oa1);

            if (oa1 != null)
            {
                Marshal.ReleaseComObject(oa1);
            }

            // assert appointment ga1 points to oa1
            ClassicAssert.AreEqual(oid_oa1, oid_ga1);
            // assert appointment ga2 points to oa1
            ClassicAssert.AreEqual(oid_oa1, oid_ga2);
            // assert appointment oa1 does not point to ga1
            ClassicAssert.AreNotEqual(gid_oa1, gid_ga1);
            // assert appointment oa1 does not point to ga2
            ClassicAssert.AreNotEqual(gid_oa1, gid_ga2);

            //Act
            Synchronizer.LoadAppointments();

            //Assert
            var f_ga1 = Synchronizer.GetGoogleAppointmentById(gid_ga1);
            var f_ga2 = Synchronizer.GetGoogleAppointmentById(gid_ga2);
            var f_oa1 = Synchronizer.GetOutlookAppointmentById(oid_oa1);

            ClassicAssert.IsNotNull(f_ga1);
            ClassicAssert.IsNotNull(f_ga2);
            ClassicAssert.IsNotNull(f_oa1);

            var f_gid_oa1 = AppointmentPropertiesUtils.GetOutlookGoogleId(f_oa1);
            var f_gid_ga1 = AppointmentPropertiesUtils.GetGoogleId(f_ga1);
            var f_gid_ga2 = AppointmentPropertiesUtils.GetGoogleId(f_ga2);
            var f_oid_ga1 = AppointmentPropertiesUtils.GetGoogleOutlookId(f_ga1);
            var f_oid_ga2 = AppointmentPropertiesUtils.GetGoogleOutlookId(f_ga2);
            var f_oid_oa1 = AppointmentPropertiesUtils.GetOutlookId(f_oa1);

            // assert appointment ga1 does not point to oa1
            ClassicAssert.AreNotEqual(f_oid_oa1, f_oid_ga1);
            // assert appointment oa1 does not point to ga1
            ClassicAssert.AreNotEqual(f_gid_oa1, f_gid_ga1);
            // assert appointment ga2 does not point to oa1
            ClassicAssert.AreNotEqual(f_oid_oa1, f_oid_ga2);
            // assert appointment oa1 does not point to ga2
            ClassicAssert.AreNotEqual(f_gid_oa1, f_gid_ga2);

            DeleteTestAppointment(f_oa1);
            DeleteTestAppointment(f_ga1);
            DeleteTestAppointment(f_ga2);

            if (f_oa1 != null)
            {
                Marshal.ReleaseComObject(f_oa1);
            }

            Log.Information($"*** TestRemoveGoogleDuplicatedAppointments_02 ***");
        }

        [Test]
        public void TestRemoveOutlookDuplicatedAppointments_01()
        {
            Log.Information($"*** TestRemoveOutlookDuplicatedAppointments_01 ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa1 = new OutlookAppointmentBuilder().BuildDefault();
            oa1.Save();

            // create new Google test appointments
            var ga1 = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa1, ref ga1);

            var oa2 = new OutlookAppointmentBuilder().BuildDefault();
            oa2.Save();
            Synchronizer.UpdateAppointment(oa2, ref ga1);

            var gid_oa1 = AppointmentPropertiesUtils.GetOutlookGoogleId(oa1);
            var gid_oa2 = AppointmentPropertiesUtils.GetOutlookGoogleId(oa2);
            var gid_ga1 = AppointmentPropertiesUtils.GetGoogleId(ga1);
            var oid_ga1 = AppointmentPropertiesUtils.GetGoogleOutlookId(ga1);
            var oid_oa1 = AppointmentPropertiesUtils.GetOutlookId(oa1);
            var oid_oa2 = AppointmentPropertiesUtils.GetOutlookId(oa2);

            if (oa1 != null)
            {
                Marshal.ReleaseComObject(oa1);
            }
            if (oa2 != null)
            {
                Marshal.ReleaseComObject(oa2);
            }

            // assert appointments oa2 and ga1 are pointing to each other
            ClassicAssert.AreEqual(gid_oa2, gid_ga1);
            ClassicAssert.AreEqual(oid_oa2, oid_ga1);
            // assert appointment oa1 also points to ga1
            ClassicAssert.AreEqual(gid_oa1, gid_ga1);
            // assert appointment ga1 does not point to oa1
            ClassicAssert.AreNotEqual(oid_ga1, oid_oa1);

            //Act
            Synchronizer.LoadAppointments();

            //Assert
            var f_ga1 = Synchronizer.GetGoogleAppointmentById(gid_ga1);
            var f_oa1 = Synchronizer.GetOutlookAppointmentById(oid_oa1);
            var f_oa2 = Synchronizer.GetOutlookAppointmentById(oid_oa2);

            ClassicAssert.IsNotNull(f_ga1);
            ClassicAssert.IsNotNull(f_oa1);
            ClassicAssert.IsNotNull(f_oa2);

            var f_gid_oa1 = AppointmentPropertiesUtils.GetOutlookGoogleId(f_oa1);
            var f_gid_oa2 = AppointmentPropertiesUtils.GetOutlookGoogleId(f_oa2);
            var f_gid_ga1 = AppointmentPropertiesUtils.GetGoogleId(f_ga1);
            var f_oid_ga1 = AppointmentPropertiesUtils.GetGoogleOutlookId(f_ga1);
            var f_oid_oa1 = AppointmentPropertiesUtils.GetOutlookId(f_oa1);
            var f_oid_oa2 = AppointmentPropertiesUtils.GetOutlookId(f_oa2);

            // assert appointments oa2 and ga1 are pointing to each other
            ClassicAssert.AreEqual(f_gid_oa2, f_gid_ga1);
            ClassicAssert.AreEqual(f_oid_oa2, f_oid_ga1);
            // assert appointment oa1 does not point to ga1
            ClassicAssert.AreNotEqual(f_oid_oa1, f_oid_ga1);
            // assert appointment oa1 does not point to ga1
            ClassicAssert.AreNotEqual(f_gid_oa1, f_gid_ga1);

            DeleteTestAppointment(f_oa1);
            DeleteTestAppointment(f_oa2);
            DeleteTestAppointment(f_ga1);

            if (f_oa1 != null)
            {
                Marshal.ReleaseComObject(f_oa1);
            }
            if (f_oa2 != null)
            {
                Marshal.ReleaseComObject(f_oa2);
            }

            Log.Information($"*** TestRemoveOutlookDuplicatedAppointments_01 ***");
        }

        [Test]
        public void TestRemoveOutlookDuplicatedAppointments_02()
        {
            Log.Information($"*** TestRemoveOutlookDuplicatedAppointments_02 ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            // create new Outlook test appointment
            var oa1 = new OutlookAppointmentBuilder().BuildDefault();
            oa1.Save();

            // create new Google test appointments
            var ga1 = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa1, ref ga1);

            var oa2 = new OutlookAppointmentBuilder().BuildDefault();
            oa2.Save();

            Synchronizer.UpdateAppointment(oa2, ref ga1);
            AppointmentPropertiesUtils.ResetGoogleOutlookId(ga1);
            ga1 = Synchronizer.SaveGoogleAppointment(ga1);

            var gid_oa1 = AppointmentPropertiesUtils.GetOutlookGoogleId(oa1);
            var gid_oa2 = AppointmentPropertiesUtils.GetOutlookGoogleId(oa2);
            var gid_ga1 = AppointmentPropertiesUtils.GetGoogleId(ga1);
            var oid_ga1 = AppointmentPropertiesUtils.GetGoogleOutlookId(ga1);
            var oid_oa1 = AppointmentPropertiesUtils.GetOutlookId(oa1);
            var oid_oa2 = AppointmentPropertiesUtils.GetOutlookId(oa2);

            if (oa1 != null)
            {
                Marshal.ReleaseComObject(oa1);
            }
            if (oa2 != null)
            {
                Marshal.ReleaseComObject(oa2);
            }

            // assert oa1 points to ga1
            ClassicAssert.AreEqual(gid_oa1, gid_ga1);
            // assert oa2 points to ga1
            ClassicAssert.AreEqual(gid_oa2, gid_ga1);
            // assert appointment ga1 does not point to oa1
            ClassicAssert.AreNotEqual(oid_ga1, oid_oa1);
            // assert appointment ga1 does not point to oa2
            ClassicAssert.AreNotEqual(oid_ga1, oid_oa2);

            //Act
            Synchronizer.LoadAppointments();

            //Assert
            var f_ga1 = Synchronizer.GetGoogleAppointmentById(gid_ga1);
            var f_oa1 = Synchronizer.GetOutlookAppointmentById(oid_oa1);
            var f_oa2 = Synchronizer.GetOutlookAppointmentById(oid_oa2);

            ClassicAssert.IsNotNull(f_ga1);
            ClassicAssert.IsNotNull(f_oa1);
            ClassicAssert.IsNotNull(f_oa2);

            var f_gid_oa1 = AppointmentPropertiesUtils.GetOutlookGoogleId(f_oa1);
            var f_gid_oa2 = AppointmentPropertiesUtils.GetOutlookGoogleId(f_oa2);
            var f_gid_ga1 = AppointmentPropertiesUtils.GetGoogleId(f_ga1);
            var f_oid_ga1 = AppointmentPropertiesUtils.GetGoogleOutlookId(f_ga1);
            var f_oid_oa1 = AppointmentPropertiesUtils.GetOutlookId(f_oa1);
            var f_oid_oa2 = AppointmentPropertiesUtils.GetOutlookId(f_oa2);

            // assert oa1 does not point to ga1
            ClassicAssert.AreNotEqual(f_gid_oa1, f_gid_ga1);
            // assert oa2 does not point to ga1
            ClassicAssert.AreNotEqual(f_gid_oa2, f_gid_ga1);
            // assert appointment ga1 does not point to oa1
            ClassicAssert.AreNotEqual(f_oid_ga1, f_oid_oa1);
            // assert appointment ga1 does not point to oa2
            ClassicAssert.AreNotEqual(f_oid_ga1, f_oid_oa2);

            DeleteTestAppointment(f_oa1);
            DeleteTestAppointment(f_oa2);
            DeleteTestAppointment(f_ga1);

            if (f_oa1 != null)
            {
                Marshal.ReleaseComObject(f_oa1);
            }
            if (f_oa2 != null)
            {
                Marshal.ReleaseComObject(f_oa2);
            }

            Log.Information($"*** TestRemoveOutlookDuplicatedAppointments_02 ***");
        }

        [Test]
        public void TestSync_Time()
        {
            Log.Information($"*** TestSync_Time ***");

            TestSync(false, false, false);

            Log.Information($"*** TestSync_Time ***");
        }


        [Test]
        public void TestSync_Day()
        {
            Log.Information($"*** TestSync_Day ***");

            TestSync(true, false, false);

            Log.Information($"*** TestSync_Day ***");
        }

        [Test]
        public void TestSync_Day_Private()
        {
            Log.Information($"*** TestSync_Day_Private ***");

            TestSync(true, false, true);

            Log.Information($"*** TestSync_Day_Private ***");
        }

        [Test]
        public void TestSync_Day_Private_Force()
        {
            Log.Information($"*** TestSync_Day_Private_Force ***");

            TestSync(true, true, true);

            Log.Information($"*** TestSync_Day_Private_Force ***");
        }

        private void TestSync(bool allDay, bool syncAppointmentsPrivate, bool privateFlag )
        {
            //Arrange
            Synchronizer.SyncOption = SyncOption.MergeOutlookWins;
            Synchronizer.SyncAppointmentsPrivate = syncAppointmentsPrivate;

            // create new appointment to sync
            Outlook.AppointmentItem oa;
            if (allDay)
                oa = new OutlookAppointmentBuilder().BuildDefault();
            else
                oa = new OutlookAppointmentBuilder().BuildAllDay();

            if (privateFlag)
                oa.Sensitivity = Outlook.OlSensitivity.olPrivate;
            else
                oa.Sensitivity = Outlook.OlSensitivity.olNormal;
            oa.Save();

            Synchronizer.SyncOption = SyncOption.OutlookToGoogleOnly;

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);
            if (!syncAppointmentsPrivate)
            {
                if (privateFlag)
                    ClassicAssert.AreEqual(AppointmentSync.PRIVATE, ga.Visibility??"default");
                else
                    ClassicAssert.AreEqual("default", ga.Visibility??"default");
            }
            else
                ClassicAssert.AreEqual(AppointmentSync.PRIVATE, ga.Visibility??"default");

            Synchronizer.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same appointment from google.
            Synchronizer.MatchAppointments();
            var match = FindMatch(oa);

            ClassicAssert.IsNotNull(match);
            ClassicAssert.IsNotNull(match.GoogleAppointment);
            ClassicAssert.IsNotNull(match.OutlookAppointment);

            var recreatedOutlookAppointment = new OutlookAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(ref match.GoogleAppointment, ref recreatedOutlookAppointment, match.GoogleAppointmentExceptions);
            Assert.That(oa, Is.Not.Null);
            ClassicAssert.IsNotNull(recreatedOutlookAppointment);
            // match recreatedOutlookAppointment with outlookAppointment

            ClassicAssert.AreEqual(oa.Subject, recreatedOutlookAppointment.Subject);
            ClassicAssert.AreEqual(oa.Start, recreatedOutlookAppointment.Start);
            ClassicAssert.AreEqual(oa.End, recreatedOutlookAppointment.End);
            ClassicAssert.AreEqual(oa.AllDayEvent, recreatedOutlookAppointment.AllDayEvent);
            //ToDo: Check other properties
            if (!syncAppointmentsPrivate)
                ClassicAssert.AreEqual(oa.Sensitivity, recreatedOutlookAppointment.Sensitivity);

            DeleteTestAppointments(match);
            recreatedOutlookAppointment.Delete();

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }
            if (recreatedOutlookAppointment != null)
            {
                Marshal.ReleaseComObject(recreatedOutlookAppointment);
            }
        }



        [Test]
        public void TestExtendedProps()
        {
            Log.Information($"*** TestExtendedProps ***");

            //Arrange
            Synchronizer.SyncOption = SyncOption.MergeOutlookWins;

            // create new appointment to sync
            var oa = new OutlookAppointmentBuilder().BuildAllDay();
            oa.Save();

            var ga = new GoogleAppointmentBuilder().Build();
            Synchronizer.UpdateAppointment(oa, ref ga);

            ClassicAssert.AreEqual(name, ga.Summary);

            // read appointment from google
            Synchronizer.MatchAppointments();
            AppointmentsMatcher.SyncAppointments(Synchronizer);

            var match = FindMatch(oa);

            ClassicAssert.IsNotNull(match);
            ClassicAssert.IsNotNull(match.GoogleAppointment);

            // get extended prop
            ClassicAssert.AreEqual(AppointmentPropertiesUtils.GetOutlookId(oa), AppointmentPropertiesUtils.GetGoogleOutlookId(match.GoogleAppointment));

            DeleteTestAppointments(match);

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            Log.Information($"*** TestExtendedProps ***");
        }

        // Two synchronized non recurring appointments
        // On Google side appointment deleted
        // Testing if sync will create Google appointment
        [Test]
        public void Test_GetGoogleAppointment()
        {
            Log.Information($"*** Test_GetGoogleAppointment ***");

            //Arrange
            var ga = new GoogleAppointmentBuilder().Build();
            ga.Summary = "GCSM Test Appointment";
            ga.Start.DateTimeDateTimeOffset = DateTime.Now;
            ga.End.DateTimeDateTimeOffset = DateTime.Now;

            ga = Synchronizer.SaveGoogleAppointment(ga);

            var ga1 = Synchronizer.GetGoogleAppointment(ga.Id);
            ClassicAssert.IsNotNull(ga1);
            ClassicAssert.AreEqual(ga.Id, ga1.Id);

            //delete test contacts
            Synchronizer.DeleteGoogleAppointment(ga);

            Log.Information($"*** Test_GetGoogleAppointment ***");
        }


        [Test]
        public void TestNotMappedRecurrence()
        {
            Log.Information($"*** TestNotMappedRecurrence ***");

            var oa = new OutlookAppointmentBuilder().Build();
            oa.Subject = $"Test";
            oa.Start = new DateTime(2009, 06, 12);
            oa.End = new DateTime(2009, 06, 13);
            oa.AllDayEvent = true;
            oa.ReminderSet = false;
            var rp = oa.GetRecurrencePattern();
            rp.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearly;
            rp.StartTime = DateTime.Parse("2009-06-12");
            rp.PatternStartDate = DateTime.Parse("2009-06-12");
            rp.PatternEndDate = DateTime.Parse("2019-06-12");
            rp.EndTime = DateTime.Parse("2009-06-13");

            try
            {
                //8 * 12 works
                rp.Interval = 9 * 12;
            }
            catch (COMException ex)
            {
                Log.Information(ex, $"Iteration");
            }

            oa.Save();
            //           var oid = AppointmentPropertiesUtils.GetOutlookId(oa);
            //           ClassicAssert.IsNotNull(oid);

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            Log.Information($"*** TestNotMappedRecurrence ***");
        }

        private void DelayBetweenSync()
        {
            //we need to wait at least 3 minutes, so next synchronization will not ignore updates
            Thread.Sleep(3 * 60 * 1000);
        }

        private void DeleteTestAppointments(AppointmentMatch match)
        {
            if (match != null)
            {
                DeleteTestAppointment(match.GoogleAppointment);
                DeleteTestAppointment(match.OutlookAppointment);
            }
        }

        private void DeleteTestAppointment(Outlook.AppointmentItem oa)
        {
            if (oa != null)
            {
                oa.Delete();
            }
        }

        private void DeleteTestAppointment(Event ga)
        {
            Synchronizer.DeleteGoogleAppointment(ga);
        }

        private void CleanAllLoadedAppointments()
        {
            for (var ii = 0; ii < Synchronizer.Appointments.Count; ii++)
            {
                var match = Synchronizer.Appointments[ii];
                if (match.OutlookAppointment != null)
                {
                    Marshal.ReleaseComObject(match.OutlookAppointment);
                }
                match.GoogleAppointment = null;
                match.GoogleAppointmentExceptions.Clear();
                match.GoogleAppointmentExceptions = null;
            }
        }

        private void DeleteTestAppointments()
        {
            Synchronizer.LoadAppointments();

            var oa = Synchronizer.OutlookAppointments.Find("[Subject] = '" + name + "'") as Outlook.AppointmentItem;
            while (oa != null)
            {
                DeleteTestAppointment(oa);
                if (oa != null)
                {
                    Marshal.ReleaseComObject(oa);
                }
                oa = Synchronizer.OutlookAppointments.Find("[Subject] = '" + name + "'") as Outlook.AppointmentItem;
            }

            foreach (var ga in Synchronizer.GoogleAppointments)
            {
                if (ga != null && ga.Summary != null && ga.Summary == name)
                {
                    DeleteTestAppointment(ga);
                }
            }
        }

        internal AppointmentMatch FindMatch(Outlook.AppointmentItem oa)
        {
            foreach (var match in Synchronizer.Appointments)
            {
                if (match.OutlookAppointment != null && match.OutlookAppointment.EntryID == oa.EntryID)
                {
                    return match;
                }
            }
            return null;
        }
    }
}