using Serilog;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    public static class OutlookExceptionExtensions
    {
        public static void ToDebugLog(this Outlook.Exception ex)
        {
            Log.Debug("** Outlook appointment exception **");

            Outlook.AppointmentItem oa = null;

            try
            {
                try
                {
                    oa = ex.AppointmentItem;
                }
                catch
                {
                }
                Log.Debug(" - Deleted: " + ex.Deleted ?? "null");
                Log.Debug(" - OriginalDate: " + ex.OriginalDate ?? "null");
                try
                {
                    Log.Debug("** Outlook appointment exception - AppointmentItem **");
                    oa?.ToDebugLog(printRecurrenceExceptions: false);
                    Log.Debug("** Outlook appointment exception - AppointmentItem **");
                }
                catch
                {
                }
            }
            finally
            {
                if (oa != null)
                {
                    Marshal.ReleaseComObject(oa);
                }
            }
            Log.Debug("** Outlook appointment exception **");
        }
    }
}
