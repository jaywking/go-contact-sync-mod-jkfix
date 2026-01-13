using NUnit.Framework;
using NUnit.Framework.Legacy;
using Serilog;
using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class OutlookTests
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

        // Testing how to undelete Outlook exceptions
        [Test]
        public void Test_Delete_Outlook_Exception()
        {
            Log.Information($"*** Test_Delete_Outlook_Exception ***");

            //Arrange

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
            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            rp = oa.GetRecurrencePattern();
            var ex = rp.GetOccurrence(new DateTime(2020, 5, 6, 15, 0, 0));
            ex.Delete();
            oa.Save();
            if (ex != null)
            {
                Marshal.ReleaseComObject(ex);
            }

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            //Act
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            rp = oa.GetRecurrencePattern();
            var currentPatternStartDate = rp.PatternStartDate;
            rp.PatternStartDate = currentPatternStartDate.AddYears(-1);
            rp.PatternStartDate = currentPatternStartDate;
            oa.Save();
            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

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

            Log.Information($"*** Test_Delete_Outlook_Exception ***");
        }

        // Testing Outlook appointment exception changing start
        [Test]
        public void Test_Change_Exception_Start()
        {
            Log.Information($"*** Test_Change_Exception_Start ***");

            //Arrange

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
            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            rp = oa.GetRecurrencePattern();
            var ex = rp.GetOccurrence(new DateTime(2020, 5, 6, 15, 0, 0));
            ex.Start = new DateTime(2020, 5, 7, 16, 0, 0);
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

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            //Act
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            rp = oa.GetRecurrencePattern();
            ex = rp.GetOccurrence(new DateTime(2020, 5, 7, 16, 0, 0));
            ex.Start = new DateTime(2020, 5, 6, 15, 0, 0);
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
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

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

            Log.Information($"*** Test_Change_Exception_Start ***");
        }

        // Testing Outlook appointment exception changing subject
        [Test]
        public void Test_Change_Exception_Subject()
        {
            Log.Information($"*** Test_Change_Exception_Subject ***");

            //Arrange

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
            var oid = AppointmentPropertiesUtils.GetOutlookId(oa);

            if (rp != null)
            {
                Marshal.ReleaseComObject(rp);
            }

            rp = oa.GetRecurrencePattern();
            var ex = rp.GetOccurrence(new DateTime(2020, 5, 6, 15, 0, 0));
            ex.Subject = "Exception";
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

            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

            //Act
            oa = Synchronizer.OutlookNameSpace.GetItemFromID(oid);
            rp = oa.GetRecurrencePattern();
            ex = rp.GetOccurrence(new DateTime(2020, 5, 6, 15, 0, 0));
            ex.Subject = "AN_OUTLOOK_TEST_APPOINTMENT";
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
            if (oa != null)
            {
                Marshal.ReleaseComObject(oa);
            }

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

            Log.Information($"*** Test_Change_Exception_Subject ***");
        }

        private void Logger_LogUpdated(string message)
        {
            Console.WriteLine(message);
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

        }

        private void DeleteTestAppointment(Outlook.AppointmentItem oa)
        {
            if (oa != null)
            {
                oa.Delete();
            }
        }
    }
}
