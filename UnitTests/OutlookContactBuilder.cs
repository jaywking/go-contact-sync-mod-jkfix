using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod.UnitTests
{
    public class OutlookContactBuilder
    {
        public Outlook.ContactItem Build()
        {
            return Synchronizer.CreateOutlookContactItem(Synchronizer.SyncContactsFolder);
        }
    }
}
