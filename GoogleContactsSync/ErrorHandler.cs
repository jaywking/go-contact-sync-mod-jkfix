using Serilog;
using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    internal static class ErrorHandler
    {

        // TODO: Write a nice error dialog, that maybe supports directly email sending as bugreport
        public static async void Handle(Exception ex)
        {
            //save user culture
            var oldCI = Thread.CurrentThread.CurrentCulture;
            //set culture to english for exception messages
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");

            Log.Error(ex.Message);
            Log.Debug(ex, "Exception");

            Log.Error("Sync failed.");

            try
            {
                SettingsForm.Instance.ShowBalloonToolTip("Error", ex.Message, ToolTipIcon.Error, 5000, true);
            }
            catch (Exception exc)
            {
                // this can fail if form was disposed or not created yet, so catch the exception - balloon is not that important to risk followup error
                Log.Error("Error showing Balloon: " + exc.Message);
            }
            //create and show error information
            using (var errorDialog = new ErrorDialog())
            {
                await errorDialog.SetErrorText(ex);
                errorDialog.ShowDialog();
            }

            //set user culture
            Thread.CurrentThread.CurrentCulture = oldCI;
            Thread.CurrentThread.CurrentUICulture = oldCI;
        }
    }
}