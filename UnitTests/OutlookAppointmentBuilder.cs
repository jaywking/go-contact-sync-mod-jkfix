using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod.UnitTests
{
    public class OutlookAppointmentBuilder
    {
        //Constants for test appointment
        public static readonly string testAppointmentSubject = "AN_OUTLOOK_TEST_APPOINTMENT";
        public Outlook.AppointmentItem Build()
        {
            return Synchronizer.CreateOutlookAppointmentItem(Synchronizer.SyncAppointmentsFolder);
        }

        public Outlook.AppointmentItem BuildDefault()
        {
            var oa = Build();
            oa.Subject = testAppointmentSubject;
            oa.Start = DateTime.Now;
            oa.End = DateTime.Now.AddHours(1);
            oa.AllDayEvent = false;
            oa.ReminderSet = false;
            return oa;
        }

        public Outlook.AppointmentItem BuildAllDay()
        {
            var oa = Build();
            oa.Subject = testAppointmentSubject;
            oa.Start = DateTime.Now;
            oa.End = DateTime.Now;
            oa.AllDayEvent = true;
            oa.ReminderSet = false;
            return oa;
        }
    }
}