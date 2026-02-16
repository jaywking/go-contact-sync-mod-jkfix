using Google.Apis.Util.Store;
using Google.Apis.Requests;
using Microsoft.Win32;
using Serilog;
using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal partial class SettingsForm : Form
    {
        #region IdleTime, see https://blog.aaronlenoir.com/2016/02/16/detect-system-idle-in-windows-applications-2/
        [DllImport("user32.dll")]
        static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        struct LASTINPUTINFO
        {
            public static readonly int SizeOf = Marshal.SizeOf(typeof(LASTINPUTINFO));

            [MarshalAs(UnmanagedType.U4)]
            public UInt32 cbSize;
            [MarshalAs(UnmanagedType.U4)]
            public UInt32 dwTime;
        }

        private TimeSpan RetrieveIdleTime()
        {
            LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = (uint)LASTINPUTINFO.SizeOf;
            GetLastInputInfo(ref lastInputInfo);

            int elapsedTicks = Environment.TickCount - (int)lastInputInfo.dwTime;

            if (elapsedTicks > 0) { return new TimeSpan(0, 0, 0, 0, elapsedTicks); }
            else { return new TimeSpan(0); }
        }
        #endregion IdleTime

        //Singleton-Object
        #region Singleton Definition

        private static volatile SettingsForm instance;
        private static readonly object syncRoot = new object();

        public static SettingsForm Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                        {
                            instance = new SettingsForm();
                        }
                    }
                }
                return instance;
            }
        }
        #endregion

        internal Synchronizer sync;
        private SyncOption syncOption;
        private DateTime lastSync;
        private bool requestClose = false;
        private bool boolShowBalloonTip = true;
        private readonly CancellationTokenSource cancellationTokenSource;
        private string prevUserName;

        public const string AppRootKey = @"Software\GoContactSyncMOD";
        public const string RegistrySyncOption = "SyncOption";
        public const string RegistryUsername = "Username";
        public const string RegistryAutoSync = "AutoSync";
        public const string RegistryAutoSyncInterval = "AutoSyncInterval";
        public const string RegistryAutoStart = "AutoStart";
        public const string RegistryReportSyncResult = "ReportSyncResult";
        public const string RegistrySyncDeletion = "SyncDeletion";
        public const string RegistryPromptDeletion = "PromptDeletion";
        public const string RegistrySyncAppointmentsMonthsInPast = "SyncAppointmentsMonthsInPast";
        public const string RegistrySyncAppointmentsMonthsInFuture = "SyncAppointmentsMonthsInFuture";
        public const string RegistrySyncAppointmentsMonthsInPastFlag = "SyncAppointmentsMonthsInPastFlag";
        public const string RegistrySyncAppointmentsMonthsInFutureFlag = "SyncAppointmentsMonthsInFutureFlag";
        public const string RegistrySyncAppointmentsTimezone = "SyncAppointmentsTimezone";
        public const string RegistrySyncAppointments = "SyncAppointments";
        public const string RegistrySyncAppointmentsForceRTF = "SyncAppointmentsForceRTF";
        public const string RegistrySyncAppointmentsPrivate = "SyncAppointmentsPrivate";
        public const string RegistrySyncOnlyIdle = "SyncOnlyIdle";
        public const string RegistrySyncReminders = "SyncReminders";
        public const string RegistryIncludePastReminders = "IncludePastReminders";
        public const string RegistrySyncContacts = "SyncContacts";
        public const string RegistrySyncContactsForceRTF = "SyncContactsForceRTF";
        public const string RegistrySyncPhotos = "SyncPhotos";
        public const string RegistryUseFileAs = "UseFileAs";
        public const string RegistryLastSync = "LastSync";
        public const string RegistrySyncContactsFolder = "SyncContactsFolder";
        public const string RegistrySyncAppointmentsFolder = "SyncAppointmentsFolder";
        public const string RegistrySyncAppointmentsGoogleFolder = "SyncAppointmentsGoogleFolder";
        public const string RegistrySyncProfile = "SyncProfile";
        private const int WaitingMinutesBeforeSync = 5;
        private readonly ProxySettingsForm _proxy = new ProxySettingsForm();

        private string syncContactsFolder = "";
        private string syncAppointmentsFolder = "";
        private string syncAppointmentsGoogleFolder = "";
        private string Timezone = "";
        private int cmbSyncProfile_PreviouslySelectedIndex = -1;

        //private string _syncProfile;
        private static string SyncProfile
        {
            get
            {
                var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
                return (regKeyAppRoot.GetValue(RegistrySyncProfile) != null) ?
                       (string)regKeyAppRoot.GetValue(RegistrySyncProfile) : null;
            }
            set
            {
                var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
                if (value != null)
                {
                    regKeyAppRoot.SetValue(RegistrySyncProfile, value);
                }
            }
        }

        private readonly string ProfileRegistry;
        private static bool OutlookFoldersLoaded = false;

        private int executing; // make this static if you want this one-caller-only to
                               // all objects instead of a single object

        private Thread syncThread;

        //register window for lock/unlock messages of workstation
        //private bool registered = false;

        private delegate void TextHandler(string text);

        private delegate void SwitchHandler(bool value);

        private delegate void IconHandler();

        private delegate DialogResult DialogHandler(string text);

        private delegate void OnTimeZoneChangesCallback(string timeZone);

        public DialogResult ShowDialog(string text)
        {
            return InvokeRequired
                ? (DialogResult)Invoke(new DialogHandler(ShowDialog), new object[] { text })
                : MessageBox.Show(this, text, Application.ProductName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
        }

        private readonly Icon IconError = Properties.Resources.sync_error;
        private readonly Icon Icon0 = Properties.Resources.sync;
        private readonly Icon Icon30 = Properties.Resources.sync_30;
        private readonly Icon Icon60 = Properties.Resources.sync_60;
        private readonly Icon Icon90 = Properties.Resources.sync_90;
        private readonly Icon Icon120 = Properties.Resources.sync_120;
        private readonly Icon Icon150 = Properties.Resources.sync_150;
        private readonly Icon Icon180 = Properties.Resources.sync_180;
        private readonly Icon Icon210 = Properties.Resources.sync_210;
        private readonly Icon Icon240 = Properties.Resources.sync_240;
        private readonly Icon Icon270 = Properties.Resources.sync_270;
        private readonly Icon Icon300 = Properties.Resources.sync_300;
        private readonly Icon Icon330 = Properties.Resources.sync_330;

        private SettingsForm()
        {
            cmbSyncProfile_PreviouslySelectedIndex = -1;

            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            cancellationTokenSource = new CancellationTokenSource();
            InitializeComponent();
            Text = Text + " - " + Application.ProductVersion;

            Program.EnableLogHandler(LogUpdatedHandler);

            Log.Information($"Started application {Application.ProductName} ({Application.ProductVersion}) on {VersionInformation.GetWindowsVersion()} and {OutlookRegistryUtils.GetOutlookVersion()}");

            var Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\GoContactSyncMOD\\";
            Log.Information($"Detailed log file created in directory: {Folder}");
            Log.Debug($"Total allocated memory at application start: {GC.GetTotalMemory(false):N0}");
            if (IsAdministrator())
            {
                Log.Information("Started with Administrator priviliges");
            }
            ContactsMatcher.NotificationReceived += new ContactsMatcher.NotificationHandler(OnNotificationReceived);
            AppointmentsMatcher.NotificationReceived += new AppointmentsMatcher.NotificationHandler(OnNotificationReceived);
            PopulateSyncOptionBox();

            //temporary remove the listener to avoid to load the settings twice, because it is set from SettingsForm.Designer.cs
            cmbSyncProfile.SelectedIndexChanged -= new EventHandler(CmbSyncProfile_SelectedIndexChanged);
            ProfileRegistry = FillSyncProfileItems() ? cmbSyncProfile.Text : null;
            LoadSettings(ProfileRegistry);

            //enable the listener
            cmbSyncProfile.SelectedIndexChanged += new EventHandler(CmbSyncProfile_SelectedIndexChanged);

            TimerSwitch(true);
            lastSyncLabel.Text = "Not synced";

            ValidateSyncButton();

            //Register Session Lock Event
            SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SystemEvents_SessionSwitch);
            //Register Power Mode Event
            SystemEvents.PowerModeChanged += new PowerModeChangedEventHandler(SystemEvents_PowerModeSwitch);
        }

        public static bool IsAdministrator()
        {
            return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void PopulateSyncOptionBox()
        {
            string str;
            for (var i = 0; i < 20; i++)
            {
                str = ((SyncOption)i).ToString();
                if (str == i.ToString())
                {
                    break;
                }

                // format (to add space before capital)
                var matches = Regex.Matches(str, "[A-Z]");
                for (var k = 0; k < matches.Count; k++)
                {
                    str = str.Replace(str[matches[k].Index].ToString(), " " + str[matches[k].Index]);
                    matches = Regex.Matches(str, "[A-Z]");
                }
                str = str.Replace("  ", " ");
                // fix start
                str = str.Substring(1);

                syncOptionBox.Items.Add(str);
            }
        }

        private void FillSyncFolderItems()
        {            

            if (InvokeRequired)
            {
                Invoke(new InvokeCallback(FillSyncFolderItems));
            }
            else
            {
                lock (syncRoot)
                {
                    if (OutlookFoldersLoaded)
                    {
                        return;
                    }

                    Log.Debug("FillSyncFolderItems - start");

                    if (contactFoldersComboBox.DataSource == null || appointmentFoldersComboBox.DataSource == null ||
                        (appointmentGoogleFoldersComboBox.DataSource == null && btSyncAppointments.Checked) ||
                        contactFoldersComboBox.Items.Count == 0 || appointmentFoldersComboBox.Items.Count == 0 ||
                        (appointmentGoogleFoldersComboBox.Items.Count == 0 && btSyncAppointments.Checked))
                    {
                        Log.Information("Loading Outlook folders...");
                        SetLastSyncText("Loading Outlook folders...");

                        contactFoldersComboBox.Visible = btSyncContactsForceRTF.Visible = SyncPhotosCheckBox.Visible = btSyncContacts.Checked;
                        labelTimezone.Visible = btMonthsPast.Visible = btMonthsFuture.Visible = btSyncAppointments.Checked;
                        appointmentFoldersComboBox.Visible = appointmentGoogleFoldersComboBox.Visible = appointmentTimezonesComboBox.Visible = btSyncAppointmentsForceRTF.Visible = btSyncAppointmentsPrivate.Visible = btSyncAppointments.Checked;
                        futureMonthInterval.Visible = btSyncAppointments.Checked && btMonthsFuture.Checked;
                        pastMonthInterval.Visible = btSyncAppointments.Checked && btMonthsPast.Checked;
                        cmbSyncProfile.Visible = true;

                        var defaultText = "    --- Select an Outlook folder ---";
                        var outlookContactFolders = new ArrayList();
                        var outlookAppointmentFolders = new ArrayList();

                        try
                        {
                            Cursor = Cursors.WaitCursor;
                            SuspendLayout();

                            contactFoldersComboBox.BeginUpdate();
                            appointmentFoldersComboBox.BeginUpdate();
                            contactFoldersComboBox.DataSource = null;
                            appointmentFoldersComboBox.DataSource = null;

                            try
                            { //Add Default Contacts Folder
                                var defaultFolder = Synchronizer.OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                                outlookContactFolders.Add(new OutlookFolder(defaultFolder.FolderPath, defaultFolder.EntryID, true));
                            }
                            catch (Exception e)
                            {
                                Log.Debug(e, "Exception");
                                Log.Warning("Error adding OlDefaultFolders.olFolderContacts: " + e.Message);
                            }

                            try
                            {//Add Default Calendar/Appointment folder
                                var defaultFolder = Synchronizer.OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                                outlookAppointmentFolders.Add(new OutlookFolder(defaultFolder.FolderPath, defaultFolder.EntryID, true));
                            }
                            catch (Exception e)
                            {
                                Log.Debug(e, "Exception");
                                Log.Warning("Error adding OlDefaultFolders.olFolderCalendar: " + e.Message);
                            }                            

                            var skipFolderScan = "1".Equals(Environment.GetEnvironmentVariable("GCSM_SKIP_OUTLOOK_FOLDER_SCAN"), StringComparison.Ordinal);
                            var defaultStoreOnly = "1".Equals(Environment.GetEnvironmentVariable("GCSM_SCAN_DEFAULT_STORE_ONLY"), StringComparison.Ordinal);
                            var skipStoreContains = Environment.GetEnvironmentVariable("GCSM_SKIP_OUTLOOK_STORE_CONTAINS");
                            var skipStoreTokens = string.IsNullOrWhiteSpace(skipStoreContains)
                                ? new string[0]
                                : skipStoreContains.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                            if (skipFolderScan)
                            {
                                Log.Warning("Skipping deep Outlook folder scan because GCSM_SKIP_OUTLOOK_FOLDER_SCAN=1. Only default folders will be listed.");
                            }
                            else
                            {
                                var folderScanStopwatch = Stopwatch.StartNew();
                                var scannedFolderCount = 0;
                                var lastProgressUpdate = DateTime.MinValue;
                                Action<string> reportFolderProgress = (folderPath) =>
                                {
                                    scannedFolderCount++;
                                    var now = DateTime.UtcNow;
                                    if ((now - lastProgressUpdate).TotalMilliseconds < 500)
                                    {
                                        return;
                                    }
                                    lastProgressUpdate = now;

                                    var elapsed = folderScanStopwatch.Elapsed;
                                    var progress = $"Loading Outlook folders... scanned {scannedFolderCount} folders in {elapsed:mm\\:ss}";
                                    if (!string.IsNullOrWhiteSpace(folderPath))
                                    {
                                        progress += $" | {folderPath}";
                                    }
                                    SetLastSyncText(progress);
                                    Application.DoEvents();
                                };

                                var folders = Synchronizer.OutlookNameSpace.Folders;
                                string defaultContactsStoreId = null;
                                if (defaultStoreOnly)
                                {
                                    try
                                    {
                                        var defaultContactsFolder = Synchronizer.OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                                        defaultContactsStoreId = defaultContactsFolder.StoreID;
                                        Log.Warning("Scanning only default Outlook store because GCSM_SCAN_DEFAULT_STORE_ONLY=1.");
                                    }
                                    catch (Exception ex)
                                    {
                                        Log.Warning($"Could not resolve default Outlook store ID, falling back to full store scan: {ex.Message}");
                                        defaultStoreOnly = false;
                                    }
                                }

                                for (var i = 1; i <= folders.Count; i++)
                                {
                                    try
                                    {
                                        var folder = folders[i];
                                        var folderPath = folder.FolderPath;

                                        if (defaultStoreOnly && !string.IsNullOrEmpty(defaultContactsStoreId) &&
                                            !string.Equals(folder.StoreID, defaultContactsStoreId, StringComparison.OrdinalIgnoreCase))
                                        {
                                            Log.Warning($"Skipping non-default Outlook store root: {folderPath}");
                                            continue;
                                        }

                                        if (skipStoreTokens.Length > 0)
                                        {
                                            var skip = false;
                                            foreach (var token in skipStoreTokens)
                                            {
                                                if (folderPath.IndexOf(token.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    Log.Warning($"Skipping Outlook store root by token '{token.Trim()}': {folderPath}");
                                                    skip = true;
                                                    break;
                                                }
                                            }
                                            if (skip)
                                            {
                                                continue;
                                            }
                                        }

                                        reportFolderProgress(folder.FolderPath);
                                        GetOutlookMAPIFolders(outlookContactFolders, outlookAppointmentFolders, folder, reportFolderProgress);

                                    }
                                    catch (System.Exception e)
                                    {
                                        Log.Debug(e, "Exception");
                                        Log.Warning("Error getting available Outlook folders: " + e.Message);
                                    }
                                }
                            }

                            if (outlookContactFolders != null)
                            {
                                outlookContactFolders.Sort();
                                outlookContactFolders.Insert(0, new OutlookFolder(defaultText, defaultText, false));
                                contactFoldersComboBox.DataSource = outlookContactFolders;
                                contactFoldersComboBox.DisplayMember = "DisplayName";
                                contactFoldersComboBox.ValueMember = "FolderID";
                            }

                            if (outlookAppointmentFolders != null)
                            {
                                outlookAppointmentFolders.Sort();
                                outlookAppointmentFolders.Insert(0, new OutlookFolder(defaultText, defaultText, false));
                                appointmentFoldersComboBox.DataSource = outlookAppointmentFolders;
                                appointmentFoldersComboBox.DisplayMember = "DisplayName";
                                appointmentFoldersComboBox.ValueMember = "FolderID";
                            }

                            contactFoldersComboBox.EndUpdate();
                            appointmentFoldersComboBox.EndUpdate();

                            contactFoldersComboBox.SelectedValue = defaultText;
                            appointmentFoldersComboBox.SelectedValue = defaultText;

                            //If user has not yet selected any folder, select one based on Outlook default folder
                            if (contactFoldersComboBox.SelectedIndex < 1)
                            {
                                foreach (OutlookFolder folder in contactFoldersComboBox.Items)
                                {
                                    if (folder.IsDefaultFolder)
                                    {
                                        contactFoldersComboBox.SelectedValue = folder.FolderID;
                                        break;
                                    }
                                }
                            }

                            //If user has not yet selected any folder, select one based on Outlook default folder
                            if (appointmentFoldersComboBox.SelectedIndex < 1)
                            {
                                foreach (OutlookFolder folder in appointmentFoldersComboBox.Items)
                                {
                                    if (folder.IsDefaultFolder)
                                    {
                                        appointmentFoldersComboBox.SelectedItem = folder;
                                        break;
                                    }
                                }
                            }

                            Log.Information("Loaded Outlook folders.");
                        }
                        catch (NotSupportedException e)
                        {
                            //Log.Warning(e.Message);
                            ErrorHandler.Handle(e);
                        }
                        catch (Exception e)
                        {
                            Log.Debug(e, "Exception");
                            Log.Warning("Error getting available Outlook and Google folders: " + e.Message);
                        }
                        finally
                        {
                            Cursor = Cursors.Default;
                            ResumeLayout();
                        }
                    }
                    LoadSettingsFolders(ProfileRegistry);

                    if ((contactFoldersComboBox.SelectedIndex == -1) && (contactFoldersComboBox.Items.Count > 0))
                    {
                        contactFoldersComboBox.SelectedIndex = 0;
                    }

                    if ((appointmentFoldersComboBox.SelectedIndex == -1) && (appointmentFoldersComboBox.Items.Count > 0))
                    {
                        appointmentFoldersComboBox.SelectedIndex = 0;
                    }

                    OutlookFoldersLoaded = true;
                    Log.Debug("FillSyncFolderItems - finish");
                }

                if (!ValidSyncContactFolders)
                {
                    Log.Debug(@"contactFoldersComboBox.SelectedIndex: " + contactFoldersComboBox.SelectedIndex);
                    Log.Debug(@"contactFoldersComboBox.Items.Count: " + contactFoldersComboBox.Items.Count);
                    throw new Exception("Outlook contact folder is not selected or invalid!");
                }

                if (!ValidSyncAppointmentFolders)
                {
                    Log.Debug($"appointmentFoldersComboBox.SelectedIndex: {appointmentFoldersComboBox.SelectedIndex}");
                    Log.Debug($"appointmentFoldersComboBox.Items.Count: {appointmentFoldersComboBox.Items.Count}");
                    Log.Debug($"appointmentGoogleFoldersComboBox.SelectedIndex: {appointmentGoogleFoldersComboBox.SelectedIndex}");
                    Log.Debug($"appointmentGoogleFoldersComboBox.Items.Count: {appointmentGoogleFoldersComboBox.Items.Count}");

                    if (appointmentGoogleFoldersComboBox.SelectedIndex <= 0)
                    {
                        throw new ApplicationException("Please select in Settings GUI a Google appointment folder to be synchronized");
                    }
                    else if (appointmentFoldersComboBox.SelectedIndex <= 0)
                    {
                        throw new ApplicationException("Please select in Settings GUI an Outlook appointment folder to be synchronized");
                    }
                    else
                    {
                        throw new ApplicationException("At least one Outlook or Google appointment folder is not selected or invalid!");
                    }
                }
            }

        }

        /// <summary>
        /// recursively scan thru all Outlook MapiFolders to search for existing Contact and Appointment folders
        /// </summary>
        /// <param name="outlookContactFolders">to be filled array with contact folders found</param>
        /// <param name="outlookAppointmentFolders">to be filled array with appointment folders found</param>
        /// <param name="folder">parent folder to scan through child folders</param>
        public static void GetOutlookMAPIFolders(ArrayList outlookContactFolders, ArrayList outlookAppointmentFolders, Outlook.MAPIFolder folder, Action<string> reportFolderProgress = null)
        {
            for (var i = 1; i <= folder.Folders.Count; i++)
            {
                Outlook.MAPIFolder mapi = null;
                try
                {
                    mapi = folder.Folders[i];
                    reportFolderProgress?.Invoke(mapi.FolderPath);
                    if (mapi.DefaultItemType == Outlook.OlItemType.olContactItem)
                    {
                        var isDefaultFolder = mapi.EntryID.Equals(Synchronizer.OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).EntryID);
                        if (!isDefaultFolder) //only add, if not Default Folder, because DefaultFolder already added before calling this method
                            outlookContactFolders.Add(new OutlookFolder(mapi.FolderPath, mapi.EntryID, isDefaultFolder));
                    }
                    if (mapi.DefaultItemType == Outlook.OlItemType.olAppointmentItem)
                    {
                        var isDefaultFolder = mapi.EntryID.Equals(Synchronizer.OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).EntryID);
                        if (!isDefaultFolder) //only add, if not Default Folder, because DefaultFolder already added before calling this method
                            outlookAppointmentFolders.Add(new OutlookFolder(mapi.FolderPath, mapi.EntryID, isDefaultFolder));
                    }

                    //Cannot be fixed by patch provided in bug #1363: https://sourceforge.net/p/googlesyncmod/bugs/1363/, because otherwise we will redundantly search ALL folders, even sub-sub-sub-... folders of Emails. Please use a Contact or Calendar root parent folder to keep your contacts and Calendars subfolders in
                    if (mapi.DefaultItemType == Outlook.OlItemType.olContactItem ||
                        mapi.DefaultItemType == Outlook.OlItemType.olAppointmentItem)
                    {
                        GetOutlookMAPIFolders(outlookContactFolders, outlookAppointmentFolders, mapi, reportFolderProgress);
                    }
                }
                catch (COMException e)
                {
                    Log.Debug(e, "Exception");
                }
                finally
                {
                    if (mapi != null)
                    {
                        Marshal.ReleaseComObject(mapi);
                    }
                }
            }
        }

        private void ClearSettings()
        {
            SetSyncOption(0);
            autoSyncCheckBox.Checked = runAtStartupCheckBox.Checked = reportSyncResultCheckBox.Checked = false;
            autoSyncInterval.Value = 2;
            _proxy.ClearSettings();
        }
        // Fill lists of sync profiles
        private bool FillSyncProfileItems()
        {
            var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);

            var vReturn = false;

            cmbSyncProfile.Items.Clear();
            cmbSyncProfile.Items.Add("[Add new profile...]");

            foreach (var subKeyName in regKeyAppRoot.GetSubKeyNames())
            {
                if (!string.IsNullOrEmpty(subKeyName))
                {
                    cmbSyncProfile.Items.Add(subKeyName);
                }
            }

            if (SyncProfile == null)
            {
                SyncProfile = "Default_" + Environment.MachineName;
            }

            if (cmbSyncProfile.Items.Count == 1)
            {
                cmbSyncProfile.Items.Add(SyncProfile);
            }
            else
            {
                vReturn = true;
            }

            cmbSyncProfile.Items.Add("[Configuration manager...]");
            cmbSyncProfile.Text = SyncProfile;

            cmbSyncProfile_PreviouslySelectedIndex = cmbSyncProfile.SelectedIndex;

            return vReturn;
        }

        private void LoadSettings(string _profile)
        {
            var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + (_profile != null ? ('\\' + _profile) : ""));

            if (regKeyAppRoot.GetValue(RegistrySyncOption) != null)
            {
                syncOption = (SyncOption)regKeyAppRoot.GetValue(RegistrySyncOption);
                SetSyncOption((int)syncOption);
            }

            if (regKeyAppRoot.GetValue(RegistryUsername) != null)
            {
                prevUserName = regKeyAppRoot.GetValue(RegistryUsername) as string;
                UserName.Text = prevUserName;
            }

            //temporary remove listener
            autoSyncCheckBox.CheckedChanged -= new EventHandler(AutoSyncCheckBox_CheckedChanged);
            btSyncContacts.CheckedChanged -= new EventHandler(BtSyncContacts_CheckedChanged);
            btSyncAppointments.CheckedChanged -= new EventHandler(BtSyncAppointments_CheckedChanged);
            SyncRemindersCheckBox.CheckedChanged -= new EventHandler(SyncReminders_CheckedChanged);

            ReadRegistryIntoCheckBox(autoSyncCheckBox, regKeyAppRoot.GetValue(RegistryAutoSync));
            ReadRegistryIntoNumber(autoSyncInterval, regKeyAppRoot.GetValue(RegistryAutoSyncInterval));
            ReadRegistryIntoCheckBox(runAtStartupCheckBox, regKeyAppRoot.GetValue(RegistryAutoStart));
            ReadRegistryIntoCheckBox(reportSyncResultCheckBox, regKeyAppRoot.GetValue(RegistryReportSyncResult));
            ReadRegistryIntoCheckBox(btSyncDelete, regKeyAppRoot.GetValue(RegistrySyncDeletion));
            ReadRegistryIntoCheckBox(btPromptDelete, regKeyAppRoot.GetValue(RegistryPromptDeletion));
            ReadRegistryIntoNumber(pastMonthInterval, regKeyAppRoot.GetValue(RegistrySyncAppointmentsMonthsInPast));
            ReadRegistryIntoNumber(futureMonthInterval, regKeyAppRoot.GetValue(RegistrySyncAppointmentsMonthsInFuture));
            ReadRegistryIntoCheckBox(btMonthsPast, regKeyAppRoot.GetValue(RegistrySyncAppointmentsMonthsInPastFlag));
            ReadRegistryIntoCheckBox(btMonthsFuture, regKeyAppRoot.GetValue(RegistrySyncAppointmentsMonthsInFutureFlag));
            if (regKeyAppRoot.GetValue(RegistrySyncAppointmentsTimezone) != null)
            {
                appointmentTimezonesComboBox.Text = regKeyAppRoot.GetValue(RegistrySyncAppointmentsTimezone) as string;
            }

            ReadRegistryIntoCheckBox(btSyncAppointments, regKeyAppRoot.GetValue(RegistrySyncAppointments));

            ReadRegistryIntoCheckBox(btSyncContacts, regKeyAppRoot.GetValue(RegistrySyncContacts));
            ReadRegistryIntoCheckBox(chkUseFileAs, regKeyAppRoot.GetValue(RegistryUseFileAs));

            ReadRegistryIntoCheckBox(btSyncContactsForceRTF, regKeyAppRoot.GetValue(RegistrySyncContactsForceRTF));
            ReadRegistryIntoCheckBox(SyncPhotosCheckBox, regKeyAppRoot.GetValue(RegistrySyncPhotos));
            ReadRegistryIntoCheckBox(btSyncAppointmentsForceRTF, regKeyAppRoot.GetValue(RegistrySyncAppointmentsForceRTF));
            ReadRegistryIntoCheckBox(btSyncAppointmentsPrivate, regKeyAppRoot.GetValue(RegistrySyncAppointmentsPrivate));
            ReadRegistryIntoCheckBox(btSyncOnlyIdle, regKeyAppRoot.GetValue(RegistrySyncOnlyIdle));
            ReadRegistryIntoCheckBox(SyncRemindersCheckBox, regKeyAppRoot.GetValue(RegistrySyncReminders));

            if (! ReadRegistryIntoCheckBox(IncludePastRemindersCheckBox, regKeyAppRoot.GetValue(RegistryIncludePastReminders)))
                IncludePastRemindersCheckBox.Checked = false;

            if (regKeyAppRoot.GetValue(RegistryLastSync) != null)
            {
                try
                {
                    lastSync = new DateTime(Convert.ToInt64(regKeyAppRoot.GetValue(RegistryLastSync)));
                    SetLastSyncText(lastSync.ToString());
                }
                catch (FormatException ex)
                {
                    Log.Warning("LastSyncDate couldn't be read from registry (" + regKeyAppRoot.GetValue(RegistryLastSync) + "): " + ex);
                }
            }

            //autoSyncCheckBox_CheckedChanged(null, null);
            BtSyncContacts_CheckedChanged(null, null);

            _proxy.LoadSettings(_profile);

            contactFoldersComboBox.Visible = btSyncContacts.Checked;
            btSyncContactsForceRTF.Visible = btSyncContacts.Checked;
            SyncPhotosCheckBox.Visible = btSyncContacts.Checked;
            appointmentFoldersComboBox.Visible = appointmentGoogleFoldersComboBox.Visible = btSyncAppointments.Checked;
            labelTimezone.Visible = btMonthsPast.Visible = btMonthsFuture.Visible = btSyncAppointments.Checked;
            appointmentTimezonesComboBox.Visible = btSyncAppointments.Checked;
            pastMonthInterval.Visible = btSyncAppointments.Checked && btMonthsPast.Checked;
            futureMonthInterval.Visible = btSyncAppointments.Checked && btMonthsFuture.Checked;
            btSyncAppointmentsForceRTF.Visible = btSyncAppointments.Checked;
            btSyncAppointmentsPrivate.Visible = SyncRemindersCheckBox.Visible = SyncRemindersCheckBox.Visible = btSyncAppointments.Checked;
            IncludePastRemindersCheckBox.Visible = btSyncAppointments.Checked && SyncRemindersCheckBox.Checked;

            //enable temporary disabled listener
            autoSyncCheckBox.CheckedChanged += new EventHandler(AutoSyncCheckBox_CheckedChanged);
            btSyncContacts.CheckedChanged += new EventHandler(BtSyncContacts_CheckedChanged);
            btSyncAppointments.CheckedChanged += new EventHandler(BtSyncAppointments_CheckedChanged);
            SyncRemindersCheckBox.CheckedChanged += new EventHandler(SyncReminders_CheckedChanged);
        }

        private static bool ReadRegistryIntoCheckBox(CheckBox checkbox, object registryEntry)
        {
            if (registryEntry == null)
                return false;

            try
            {
                checkbox.Checked = Convert.ToBoolean(registryEntry);
                return true;
            }
            catch (FormatException ex)
            {
                Log.Warning(checkbox.Name + " couldn't be read from registry (" + registryEntry + "), was kept at default (" + checkbox.Checked + "): " + ex);
                return false;
            }
        }

        private static void ReadRegistryIntoNumber(NumericUpDown numericUpDown, object registryEntry)
        {
            if (registryEntry != null)
            {
                var interval = Convert.ToDecimal(registryEntry);
                if (interval < numericUpDown.Minimum)
                {
                    numericUpDown.Value = numericUpDown.Minimum;
                    Log.Warning(numericUpDown.Name + " read from registry was below range (" + interval + "), was set to minimum (" + numericUpDown.Minimum + ")");
                }
                else if (interval > numericUpDown.Maximum)
                {
                    numericUpDown.Value = numericUpDown.Maximum;
                    Log.Warning(numericUpDown.Name + " read from registry was above range (" + interval + "), was set to maximum (" + numericUpDown.Maximum + ")");
                }
                else
                {
                    numericUpDown.Value = interval;
                }
            }
        }

        private void LoadSettingsFolders(string _profile)
        {
            Log.Debug("Loading settings folders...");

            var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + (_profile != null ? ('\\' + _profile) : ""));

            var regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncContactsFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {
                foreach (OutlookFolder i in contactFoldersComboBox.Items)
                {
                    if (i.FolderID == regKeyValueStr)
                    {
                        contactFoldersComboBox.SelectedValue = regKeyValueStr;
                        break;
                    }
                }
            }

            regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncAppointmentsFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {
                foreach (OutlookFolder i in appointmentFoldersComboBox.Items)
                {
                    if (i.FolderID == regKeyValueStr)
                    {
                        appointmentFoldersComboBox.SelectedValue = regKeyValueStr;
                        break;
                    }
                }
            }
            
            if (appointmentGoogleFoldersComboBox.DataSource == null)
            {
                LoadAppointmentGoogleFoldersComboBox();
            }

            regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncAppointmentsGoogleFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {              
                appointmentGoogleFoldersComboBox.SelectedValue = regKeyValueStr;
                if ((appointmentGoogleFoldersComboBox.SelectedIndex == -1) && (appointmentGoogleFoldersComboBox.Items.Count > 0))
                {
                    appointmentGoogleFoldersComboBox.SelectedIndex = 0;
                }
            }

            Log.Debug("Loaded settings folders...");
        }

        private void SaveSettings()
        {
            SaveSettings(cmbSyncProfile.Text);
        }

        private void SaveSettings(string profile)
        {
            if (!string.IsNullOrEmpty(profile))
            {
                SyncProfile = cmbSyncProfile.Text;
                var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + profile);
                regKeyAppRoot.SetValue(RegistrySyncOption, (int)syncOption);

                if (!string.IsNullOrEmpty(UserName.Text))
                {
                    regKeyAppRoot.SetValue(RegistryUsername, UserName.Text);
                }
                regKeyAppRoot.SetValue(RegistryAutoSync, autoSyncCheckBox.Checked.ToString());
                regKeyAppRoot.SetValue(RegistryAutoSyncInterval, autoSyncInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistryAutoStart, runAtStartupCheckBox.Checked);
                regKeyAppRoot.SetValue(RegistryReportSyncResult, reportSyncResultCheckBox.Checked);
                regKeyAppRoot.SetValue(RegistrySyncDeletion, btSyncDelete.Checked);
                regKeyAppRoot.SetValue(RegistryPromptDeletion, btPromptDelete.Checked);
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsMonthsInPast, pastMonthInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsMonthsInFuture, futureMonthInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsMonthsInPastFlag, btMonthsPast.Checked.ToString());
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsMonthsInFutureFlag, btMonthsFuture.Checked.ToString());
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsTimezone, appointmentTimezonesComboBox.Text);
                regKeyAppRoot.SetValue(RegistrySyncAppointments, btSyncAppointments.Checked);
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsForceRTF, btSyncAppointmentsForceRTF.Checked);
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsPrivate, btSyncAppointmentsPrivate.Checked);
                regKeyAppRoot.SetValue(RegistrySyncOnlyIdle, btSyncOnlyIdle.Checked);
                regKeyAppRoot.SetValue(RegistrySyncReminders, SyncRemindersCheckBox.Checked);                
                regKeyAppRoot.SetValue(RegistryIncludePastReminders, IncludePastRemindersCheckBox.Checked);                
                regKeyAppRoot.SetValue(RegistrySyncContacts, btSyncContacts.Checked);
                regKeyAppRoot.SetValue(RegistrySyncContactsForceRTF, btSyncContactsForceRTF.Checked);
                regKeyAppRoot.SetValue(RegistrySyncPhotos, SyncPhotosCheckBox.Checked);
                regKeyAppRoot.SetValue(RegistryUseFileAs, chkUseFileAs.Checked);
                regKeyAppRoot.SetValue(RegistryLastSync, lastSync.Ticks);

                _proxy.SaveSettings(cmbSyncProfile.Text);
            }
        }

        private bool ValidSyncFolders => ValidSyncContactFolders && ValidSyncAppointmentFolders;

        private bool ValidSyncContactFolders => (contactFoldersComboBox.SelectedIndex >= 1 && contactFoldersComboBox.SelectedIndex < contactFoldersComboBox.Items.Count)
                                                || !btSyncContacts.Checked;

        private bool ValidSyncAppointmentFolders => ValidSyncOutlookAppointmentFolders && ValidSyncGoogleAppointmentFolders || !btSyncAppointments.Checked;

        private bool ValidSyncOutlookAppointmentFolders => appointmentFoldersComboBox.SelectedIndex >= 1 && appointmentFoldersComboBox.SelectedIndex < appointmentFoldersComboBox.Items.Count;
        private bool ValidSyncGoogleAppointmentFolders => appointmentGoogleFoldersComboBox.SelectedIndex >= 1 && appointmentGoogleFoldersComboBox.SelectedIndex <appointmentGoogleFoldersComboBox.Items.Count;

        private bool ValidUserName
        {
            get
            {
                var userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);                

                SetBgColor(UserName, userNameIsValid);
                
                if (!userNameIsValid)
                {
                    toolTip.SetToolTip(UserName, "User is of wrong format, should be full Google Mail address, e.g. user@gmail.com");
                }
                else
                {
                    toolTip.SetToolTip(UserName, string.Empty);
                }

                return userNameIsValid;
            }
        }

        private bool ValidProfile
        {
            get
            {
                var syncProfileIsValid = cmbSyncProfile.SelectedIndex > 0 && cmbSyncProfile.SelectedIndex < cmbSyncProfile.Items.Count - 1;
                SetBgColor(cmbSyncProfile, syncProfileIsValid);

                return syncProfileIsValid;
            }

        }

        private static void SetBgColor(Control box, bool isValid)
        {
            box.BackColor = !isValid ? Color.LightPink : (box is ListBox)? ListBox.DefaultBackColor: (box is CheckBox) ? CheckBox.DefaultBackColor: Color.LightGreen;
        }


        private void SyncButton_Click(object sender, EventArgs e)
        {
            Sync();
        }

        private async void Sync()
        {
            try
            {
                if (syncThread != null && syncThread.IsAlive)
                {
                    TimerSwitch(false);
                    Log.Information("The sync thread is already running and alive, waiting for sync thread to finish");
                    await Task.Delay(15 * 3600 * 1000, this.cancellationTokenSource.Token); //Wait for 15 minutes
                    if (syncThread != null && syncThread.IsAlive) //Check if still alive after 15 minutes
                    {
                        Log.Information("The sync thread is still running and alive after waiting, aborting the running thread and try to start a new thread after next sync interval");
                        //cancel current and continue starting a new sync thread
                        CancelButton_Click(this, EventArgs.Empty);
                    }

                    return; //return and wait for next sync interval
                }

                if (!ValidUserName)
                {
                    //return;
                    throw new Exception("E-Mail address is incomplete or incorrect - Maybe a typo...");
                }

                if (!ValidProfile)
                {
                    //return;
                    throw new Exception("Please select a Sync Profile...");
                }

                FillSyncFolderItems();

                

                //IconTimerSwitch(true);
               
                
                var starter = new ThreadStart(Sync_ThreadStarter);
                syncThread = new Thread(starter)
                {
                    CurrentCulture = CultureInfo.CreateSpecificCulture("en-US"),
                    CurrentUICulture = new CultureInfo("en-US")
                };
                syncThread.Start();

                //if new version on sourceforge.net website than print an information to the log
                Log.Debug("Check version from Sync");
                CheckVersion();

                // wait for thread to start
                for (var i = 0; !syncThread.IsAlive && i < 10; i++)
                {
                    Thread.Sleep(1000);//Do nothing, until the thread was started, but only wait maximum 10 seconds
                }
            }
            catch (Exception ex)
            {
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(ex);
            }
        }

        [STAThread]
        private async void Sync_ThreadStarter()
        {
            //==>Instead of lock, use Interlocked to exit the code, if already another thread is calling the same
            var won = false;

            try
            {
                won = Interlocked.CompareExchange(ref executing, 1, 0) == 0;
                if (won)
                {
                    TimerSwitch(false);

                    //if the contacts folder has changed ==> Reset matches (to not delete contacts on the one or other side)                
                    var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + SyncProfile);
                    var oldSyncContactsFolder = regKeyAppRoot.GetValue(RegistrySyncContactsFolder) as string;
                    var oldSyncAppointmentsFolder = regKeyAppRoot.GetValue(RegistrySyncAppointmentsFolder) as string;
                    var oldSyncAppointmentsGoogleFolder = regKeyAppRoot.GetValue(RegistrySyncAppointmentsGoogleFolder) as string;

                    //only reset contacts if ContactsFolder changed
                    //and only reset appointments, if either OutlookAppointmentsFolder changed (without changing Google at the same time) or GoogleAppointmentsFolder changed (without changing Outlook at the same time) (not chosen before means not changed)
                    var syncContacts = !string.IsNullOrEmpty(oldSyncContactsFolder) && !oldSyncContactsFolder.Equals(syncContactsFolder) && btSyncContacts.Checked;
                    var syncAppointments = !string.IsNullOrEmpty(oldSyncAppointmentsFolder) && !oldSyncAppointmentsFolder.Equals(syncAppointmentsFolder) && btSyncAppointments.Checked;
                    var syncGoogleAppointments = !string.IsNullOrEmpty(syncAppointmentsGoogleFolder) && !syncAppointmentsGoogleFolder.Equals(oldSyncAppointmentsGoogleFolder) && btSyncAppointments.Checked;
                    if (syncContacts || (syncAppointments && !syncGoogleAppointments) || (!syncAppointments && syncGoogleAppointments))
                    {
                        var response = ShowDialog("One or more Outlook or Google folder(s) changed, do you want to reset the matches first?");
                        switch (response)
                        {
                            case DialogResult.Yes:
                                var r = await ResetMatches(syncContacts,syncAppointments||syncGoogleAppointments);
                                if (!r)
                                {
                                    throw new Exception("Reset required but cancelled by user");
                                }
                                break;
                            case DialogResult.Cancel:
                                return;
                                //default: //or DialogResult.No: Continue without Reset matches
                        }
                    }

                    //Then save the Contacts Folders used at last sync
                    if (btSyncContacts.Checked)
                    {
                        regKeyAppRoot.SetValue(RegistrySyncContactsFolder, syncContactsFolder);
                    }

                    if (btSyncAppointments.Checked)
                    {
                        regKeyAppRoot.SetValue(RegistrySyncAppointmentsFolder, syncAppointmentsFolder);
                        if (string.IsNullOrEmpty(syncAppointmentsGoogleFolder) && !string.IsNullOrEmpty(oldSyncAppointmentsGoogleFolder))
                        {
                            syncAppointmentsGoogleFolder = oldSyncAppointmentsGoogleFolder;
                        }

                        if (!string.IsNullOrEmpty(syncAppointmentsGoogleFolder))
                        {
                            regKeyAppRoot.SetValue(RegistrySyncAppointmentsGoogleFolder, syncAppointmentsGoogleFolder);
                        }
                    }

                    SetLastSyncText("Syncing...");
                    notifyIcon.Text = Application.ProductName + "\nSyncing...";
                    IconTimerSwitch(true);

                    SetFormEnabled(false);

                    if (sync == null)
                    {
                        sync = new Synchronizer();
                        sync.DuplicatesFound += new Synchronizer.DuplicatesFoundHandler(OnDuplicatesFound);
                        sync.ErrorEncountered += new Synchronizer.ErrorNotificationHandler(OnErrorEncountered);
                        sync.TimeZoneChanges += new Synchronizer.TimeZoneNotificationHandler(OnTimeZoneChanges);
                    }

                    /* Log Logger.ClearLog(); */
                    SetSyncConsoleText("");
                    Log.Information("Sync started (" + SyncProfile + ").");
                    /*SetSyncConsoleText(Logger.GetText());*/
                    Synchronizer.SyncProfile = SyncProfile;
                    Synchronizer.SyncContactsFolder = syncContactsFolder;
                    Synchronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
                    Synchronizer.SyncAppointmentsGoogleFolder = syncAppointmentsGoogleFolder;
                    Synchronizer.RestrictMonthsInPast = btMonthsPast.Checked;
                    Synchronizer.MonthsInPast = Convert.ToUInt16(pastMonthInterval.Value);
                    Synchronizer.RestrictMonthsInFuture = btMonthsFuture.Checked;
                    Synchronizer.MonthsInFuture = Convert.ToInt16(futureMonthInterval.Value);
                    
                    Synchronizer.SyncAppointmentsPrivate = btSyncAppointmentsPrivate.Checked;
                    Synchronizer.SyncReminders = SyncRemindersCheckBox.Checked;
                    Synchronizer.IncludePastReminders = IncludePastRemindersCheckBox.Checked;
                    Synchronizer.Timezone = Timezone;

                    sync.SyncOption = syncOption;
                    sync.SyncDelete = btSyncDelete.Checked;
                    sync.PromptDelete = btPromptDelete.Checked && btSyncDelete.Checked;
                    sync.UseFileAs = chkUseFileAs.Checked;
                    sync.SyncContacts = btSyncContacts.Checked;
                    sync.SyncAppointments = btSyncAppointments.Checked;
                    Synchronizer.SyncAppointmentsForceRTF = btSyncAppointmentsForceRTF.Checked;
                    Synchronizer.SyncContactsForceRTF = btSyncContactsForceRTF.Checked;
                    Synchronizer.SyncPhotos = SyncPhotosCheckBox.Checked;

                    if (!sync.SyncContacts && !sync.SyncAppointments)
                    {
                        SetLastSyncText("Sync failed.");
                        notifyIcon.Text = Application.ProductName + "\nSync failed";

                        var messageText = "Neither contacts nor appointments are switched on for syncing. Please choose at least one option. Sync aborted!";
                        Log.Error(messageText);
                        ShowForm();
                        ShowBalloonToolTip("Error", messageText, ToolTipIcon.Error, 5000, true);
                        return;
                    }

                    sync.LoginToGoogle(UserName.Text);
                    sync.LoginToOutlook();

                    sync.Sync();

                    lastSync = DateTime.Now;
                    SetLastSyncText("Last synced at " + lastSync.ToString());

                    var message = $"Sync complete. Synced: {sync.SyncedCount} out of {sync.TotalCount}. Deleted: {sync.DeletedCount}. Skipped: {sync.SkippedCount}. Errors: {sync.ErrorCount}.";
                    Log.Information(message);

                    if (reportSyncResultCheckBox.Checked)
                    {
                        var icon = sync.ErrorCount > 0 ? ToolTipIcon.Error : sync.SkippedCount > 0 ? ToolTipIcon.Warning : ToolTipIcon.Info;

                        ShowBalloonToolTip(Application.ProductName,
                            $"{DateTime.Now}. {message}",
                            icon,
                            5000, false);

                    }
                    var s = DateTime.Now.ToString("dd.MM. HH:mm");
                    var toolTip = $"{Application.ProductName}\nLast sync: {s}";
                    if (sync.ErrorCount + sync.SkippedCount > 0)
                    {
                        toolTip += $"\nWarnings: {sync.ErrorCount + sync.SkippedCount}.";
                    }

                    if (toolTip.Length >= 64)
                    {
                        toolTip = toolTip.Substring(0, 63);
                    }

                    notifyIcon.Text = toolTip;
                }
            }
            catch (Google.GoogleApiException ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                if (ex.InnerException is WebException)
                {
                    var message = "Cannot connect to Google, please check for available internet connection and proxy settings if applicable: " + ex.Message + "\r\n" + ex.InnerException.Message + "\r\n" + ex.HelpLink;
                    Log.Warning(message);
                    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }
            //catch (Google.GData.Client.InvalidCredentialsException) //ToDo: Check counterpart in new Google Api
            //{
            //    SetLastSyncText("Sync failed.");
            //    notifyIcon.Text = Application.ProductName + "\nSync failed";

            //    var message = "The credentials (Google Account username and/or password) are invalid, please correct them in the settings form before you sync again";
            //    Log.Error(message);
            //    ShowForm();
            //    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
            //}
            catch (Exception ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";
                Log.Debug(ex, "Exception");
                if (ex is COMException)
                {
                    var message = "Outlook exception, please assure that Outlook is running and not closed when syncing";
                    Log.Warning(message + ": " + ex.Message + "\r\n" + ex.StackTrace);
                    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }
            finally
            {
                if (won)
                {
                    Interlocked.Exchange(ref executing, 0);
                    lastSync = DateTime.Now;
                    TimerSwitch(true);
                    SetFormEnabled(true);
                    if (sync != null)
                    {
                        sync.LogoffOutlook();
                        sync.LogoffGoogle();
                        sync = null;
                    }
                    IconTimerSwitch(false);
                }
            }
        }

        public void ShowBalloonToolTip(string title, string message, ToolTipIcon icon, int timeout, bool error)
        {
            //if user is active on workstation
            if (boolShowBalloonTip)
            {
                notifyIcon.BalloonTipTitle = title;
                notifyIcon.BalloonTipText = message;
                notifyIcon.BalloonTipIcon = icon;
                notifyIcon.ShowBalloonTip(timeout);
            }

            var iconText = title + ": " + message;
            if (!string.IsNullOrEmpty(iconText))
            {
                notifyIcon.Text = iconText.Substring(0, iconText.Length >= 63 ? 63 : iconText.Length);
            }

            if (error)
            {
                notifyIcon.Icon = IconError;
            }
        }

        private void LogUpdatedHandler(string Message)
        {
            AppendSyncConsoleText(Message);
        }

        private void OnErrorEncountered(string title, Exception ex)
        {
            // do not show ErrorHandler, as there may be multiple exceptions that would nag the user
            Log.Error(ex.Message);
            Log.Debug(ex, "Exception");
            var message = $"Error Saving Person: {ex.Message}.\nPlease report complete ErrorMessage from Log to the Tracker\nat https://sourceforge.net/tracker/?group_id=369321";
            ShowBalloonToolTip(title, message, ToolTipIcon.Error, 5000, true);
        }

        private void OnTimeZoneChanges(string timeZone)
        {
            if (appointmentTimezonesComboBox.InvokeRequired)
            {
                var d = new OnTimeZoneChangesCallback(OnTimeZoneChanges);
                Invoke(d, new object[] { timeZone });
            }
            else
            {
                appointmentTimezonesComboBox.Text = timeZone;
            }
            Timezone = timeZone;
            Synchronizer.Timezone = timeZone;
        }

        private void OnDuplicatesFound(string title, string message)
        {
            Log.Warning(message);
            ShowBalloonToolTip(title, message, ToolTipIcon.Warning, 5000, true);
        }

        private void OnNotificationReceived(string message)
        {
            SetLastSyncText(message);
        }

        public void SetFormEnabled(bool enabled)
        {
            if (InvokeRequired)
            {
                var h = new SwitchHandler(SetFormEnabled);
                Invoke(h, new object[] { enabled });
            }
            else
            {
                resetMatchesLinkLabel.Enabled = enabled;
                settingsGroupBox.Enabled = enabled;
                syncButton.Enabled = enabled;
                cancelButton.Enabled = !enabled;
            }
        }

        public void SetLastSyncText(string text)
        {
            if (InvokeRequired)
            {
                var h = new TextHandler(SetLastSyncText);
                Invoke(h, new object[] { text });
            }
            else
            {
                lastSyncLabel.Text = text;
            }
        }

        public void SetSyncConsoleText(string text)
        {
            if (InvokeRequired)
            {
                var h = new TextHandler(SetSyncConsoleText);
                Invoke(h, new object[] { text });
            }
            else
            {
                syncConsole.Text = text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }
        }

        public void AppendSyncConsoleText(string text)
        {
            if (InvokeRequired)
            {
                var h = new TextHandler(AppendSyncConsoleText);
                Invoke(h, new object[] { text });
            }
            else
            {
                syncConsole.Text += text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }
        }

        public void TimerSwitch(bool value)
        {
            if (InvokeRequired)
            {
                var h = new SwitchHandler(TimerSwitch);
                Invoke(h, new object[] { value });
            }
            else
            {
                //If PC resumes or unlocks or is started, give him 5 minutes to recover everything before the sync starts
                if (lastSync < DateTime.Now.AddMinutes(WaitingMinutesBeforeSync) - new TimeSpan((int)autoSyncInterval.Value, 0, 0))
                    lastSync = DateTime.Now.AddMinutes(WaitingMinutesBeforeSync) - new TimeSpan((int)autoSyncInterval.Value, 0, 0);

                var v = autoSyncCheckBox.Checked && value;

                autoSyncInterval.Enabled = v;
                syncTimer.Enabled = v;
                btSyncOnlyIdle.Enabled = v;
                nextSyncLabel.Visible = v;
            }
        }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                //System shutdown
                case NativeMethods.WM_QUERYENDSESSION:
                    requestClose = true;
                    break;
                default:
                    break;
            }
            //Show Window from Tray
            if (m.Msg == NativeMethods.WM_GCSM_SHOWME)
            {
                Log.Debug("WM_GCSM_SHOWME");
                ShowForm();
            }

            base.WndProc(ref m);
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!requestClose)
            {
                SaveSettings();
                e.Cancel = true;
            }
            HideForm();
        }

        private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                cancellationTokenSource.Cancel();

                if (sync != null)
                {
                    sync.LogoffOutlook();
                }

                Log.Information("Closed application.");
                Log.CloseAndFlush();

                SaveSettings();

                //unregister event handler
                SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
                SystemEvents.PowerModeChanged -= SystemEvents_PowerModeSwitch;

                Program.DisableLogHandler(LogUpdatedHandler);

                notifyIcon.Dispose();
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }
        }

        private void SyncOptionBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Application.DoEvents();
                var index = syncOptionBox.SelectedIndex;
                if (index == -1)
                {
                    return;
                }

                SetSyncOption(index);
            }
            catch (Exception ex)
            {
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(ex);
            }
        }

        private void SetSyncOption(int index)
        {
            syncOption = (SyncOption)index;
            for (var i = 0; i < syncOptionBox.Items.Count; i++)
            {
                if (i == index)
                {
                    syncOptionBox.SetItemCheckState(i, CheckState.Checked);
                }
                else
                {
                    syncOptionBox.SetItemCheckState(i, CheckState.Unchecked);
                }
            }
        }

        private void SettingsForm_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
            }
        }

        private void NotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ShowForm();
        }

        private void AutoSyncCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            lastSync = DateTime.Now.AddMinutes(WaitingMinutesBeforeSync) - new TimeSpan((int)autoSyncInterval.Value, 0, 0);
            TimerSwitch(true);
        }

        private void SyncTimer_Tick(object sender, EventArgs e)
        {
            var syncTime = DateTime.Now - lastSync;
            var limit = new TimeSpan((int)autoSyncInterval.Value, 0, 0);

            string str = "Next sync in";
            TimeSpan diff = TimeSpan.Zero;
            TimeSpan diff2 = TimeSpan.Zero;


            if (btSyncOnlyIdle.Checked && RetrieveIdleTime() < TimeSpan.FromMinutes(WaitingMinutesBeforeSync))
            //only sync, if idleTime longer than 5 minutes           
            {
                diff = TimeSpan.FromMinutes(WaitingMinutesBeforeSync).Subtract(RetrieveIdleTime());
            }

            if (syncTime < limit)
            {
                diff2 = limit.Subtract(syncTime);
            }

            if (diff2 > diff)
                diff = diff2; //use the biggest difference to countdown

            if (diff != TimeSpan.Zero)
            {
                if (diff.Hours != 0)
                    str += " " + diff.Hours + " h";
                if (diff.Minutes != 0 || diff.Hours != 0)
                    str += " " + diff.Minutes + " min";
                if (diff.Seconds != 0)
                    str += " " + diff.Seconds + " s";

                nextSyncLabel.Text = str;
            }
            else
            {
                Sync();
            }
        }

        private async void ResetMatchesLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // force deactivation to show up
            Application.DoEvents();
            try
            {
                cancelButton.Enabled = false; //Cancel is only working for sync currently, not for reset
                await ResetMatches(btSyncContacts.Checked, btSyncAppointments.Checked);
            }
            catch (Exception ex)
            {
                SetLastSyncText("Reset Matches failed");
                Log.Error("Reset Matches failed");
                ErrorHandler.Handle(ex);
            }
            finally
            {
                lastSync = DateTime.Now;
                TimerSwitch(true);
                SetFormEnabled(true);
                hideButton.Enabled = true;
                if (sync != null)
                {
                    sync.LogoffOutlook();
                    sync.LogoffGoogle();
                    sync = null;
                }
            }
        }

        private async Task<bool> ResetMatches(bool syncContacts, bool syncAppointments)
        {
            TimerSwitch(false);

            SetLastSyncText("Resetting matches...");
            notifyIcon.Text = Application.ProductName + "\nResetting matches...";

            SetFormEnabled(false);

            if (sync == null)
            {
                sync = new Synchronizer();
            }

            /* Log Logger.ClearLog(); */
            SetSyncConsoleText("");
            Log.Information("Reset Matches started  (" + SyncProfile + ").");

            sync.SyncContacts = syncContacts;
            sync.SyncAppointments = syncAppointments;

            Synchronizer.SyncContactsFolder = syncContactsFolder;
            Synchronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
            Synchronizer.SyncAppointmentsGoogleFolder = syncAppointmentsGoogleFolder;
            Synchronizer.SyncProfile = SyncProfile;

            sync.LoginToGoogle(UserName.Text);
            sync.LoginToOutlook();

            if (sync.SyncAppointments)
            {
                bool deleteOutlookAppointments;
                switch (ShowDialog("Do you want to delete all Outlook Calendar entries?"))
                {
                    case DialogResult.Yes: deleteOutlookAppointments = true; break;
                    case DialogResult.No: deleteOutlookAppointments = false; break;
                    default: return false;
                }

                bool deleteGoogleAppointments;
                switch (ShowDialog("Do you want to delete all Google Calendar entries?"))
                {
                    case DialogResult.Yes: deleteGoogleAppointments = true; break;
                    case DialogResult.No: deleteGoogleAppointments = false; break;
                    default: return false;
                }

                Log.Information("Resetting Google appointment matches...");
                try
                {
                    await sync.ResetGoogleAppointmentMatches(deleteGoogleAppointments, cancellationTokenSource.Token);
                    sync.LoadAppointments();
                    sync.ResetOutlookAppointmentMatches(deleteOutlookAppointments);
                }
                catch (TaskCanceledException)
                {
                    Log.Information("Task cancelled by user.");
                    sync.LoadAppointments();
                }
            }

            if (sync.SyncContacts)
            {
                sync.LoadContacts();
                sync.ResetContactMatches();
            }

            lastSync = DateTime.Now;
            SetLastSyncText("Matches reset at " + lastSync.ToString());
            Log.Information("Matches reset.");

            return true;
        }

        public delegate DialogResult InvokeConflict(ConflictResolverForm conflictResolverForm);

        public DialogResult ShowConflictDialog(ConflictResolverForm conflictResolverForm)
        {
            if (InvokeRequired)
            {
                return (DialogResult)Invoke(new InvokeConflict(ShowConflictDialog), new object[] { conflictResolverForm });
            }
            else
            {
                var res = conflictResolverForm.ShowDialog(this);
                notifyIcon.Icon = Icon0;
                return res;
            }
        }

        public delegate DialogResult InvokeDeleteTooManyPropertiesForm(DeleteTooManyPropertiesForm form);

        public DialogResult ShowDeleteTooManyPropertiesForm(DeleteTooManyPropertiesForm form)
        {
            return InvokeRequired
                ? (DialogResult)Invoke(new InvokeDeleteTooManyPropertiesForm(ShowDeleteTooManyPropertiesForm), new object[] { form })
                : form.ShowDialog(this);
        }

        public delegate DialogResult InvokeDeleteTooBigPropertiesForm(DeleteTooBigPropertiesForm form);

        public DialogResult ShowDeleteTooBigPropertiesForm(DeleteTooBigPropertiesForm form)
        {
            return InvokeRequired
                ? (DialogResult)Invoke(new InvokeDeleteTooBigPropertiesForm(ShowDeleteTooBigPropertiesForm), new object[] { form })
                : form.ShowDialog(this);
        }

        public delegate DialogResult InvokeDeleteDuplicatedPropertiesForm(DeleteDuplicatedPropertiesForm form);

        public DialogResult ShowDeleteDuplicatedPropertiesForm(DeleteDuplicatedPropertiesForm form)
        {
            return InvokeRequired
                ? (DialogResult)Invoke(new InvokeDeleteDuplicatedPropertiesForm(ShowDeleteDuplicatedPropertiesForm), new object[] { form })
                : form.ShowDialog(this);
        }

        private delegate void InvokeCallback();

        private void ShowForm()
        {
            if (InvokeRequired)
            {
                Invoke(new InvokeCallback(ShowForm));
            }
            else
            {
                var oldState = WindowState;

                Show();
                Activate();
                WindowState = FormWindowState.Normal;

                using (var filter = new OleMessageFilter())
                {
                    try
                    {
                        FillSyncFolderItems();
                    }
                    catch (ApplicationException ex)
                    {
                        TimerSwitch(false);
                        Log.Error(ex.Message);
                    }
                    catch (Exception ex)
                    {
                        TimerSwitch(false);
                        ErrorHandler.Handle(ex);
                    }
                }
                               

                if (oldState != WindowState)
                {
                    Log.Debug("Check version from ShowForm, oldState: " + oldState + ", currentState: " + WindowState);
                    CheckVersion();
                }
            }
        }

        private async void CheckVersion()
        {
            if (!NewVersionLinkLabel.Visible)
            {//Only check once, if new version is available

                try
                {
                    Cursor = Cursors.WaitCursor;
                    SuspendLayout();
                    //check for new version
                    if (NewVersionLinkLabel.LinkColor != Color.Red && await VersionInformation.IsNewVersionAvailable(cancellationTokenSource.Token))
                    {
                        NewVersionLinkLabel.Visible = true;
                        NewVersionLinkLabel.LinkColor = Color.Red;
                        NewVersionLinkLabel.Text = "New Version of GCSM available on sf.net!";
                        notifyIcon.BalloonTipClicked += NotifyIcon_BalloonTipClickedDownloadNewVersion;
                        ShowBalloonToolTip("New version available", "Click here to download", ToolTipIcon.Info, 20000, false);
                    }
                    NewVersionLinkLabel.Visible = true;
                }
                finally
                {
                    Cursor = Cursors.Default;
                    ResumeLayout();
                }
            }
        }

        private void NotifyIcon_BalloonTipClickedDownloadNewVersion(object sender, EventArgs e)
        {
            Process.Start("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
            notifyIcon.BalloonTipClicked -= NotifyIcon_BalloonTipClickedDownloadNewVersion;
        }

        private void HideForm()
        {
            WindowState = FormWindowState.Minimized;
            Hide();
        }

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm();
            Activate();
        }

        private void ToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            HideForm();
        }

        private void ToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            requestClose = true;
            Close();
        }

        private void ToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            using (var about = new AboutBox())
            {
                about.Show();
            }
        }

        private void ToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Sync();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            var showWindowOnStart = !"0".Equals(Environment.GetEnvironmentVariable("GCSM_SHOW_WINDOW_ON_START"), StringComparison.Ordinal);
            if (string.IsNullOrEmpty(UserName.Text) ||
                string.IsNullOrEmpty(cmbSyncProfile.Text))
            {
                // this is the first load, show form
                ShowForm();
                UserName.Focus();
                ShowBalloonToolTip(Application.ProductName,
                        "Application started and visible in your PC's system tray, click on this balloon or the icon below to open the settings form and enter your Google credentials there.",
                        ToolTipIcon.Info,
                        5000, false);
            }
            else
            {
                if (showWindowOnStart)
                {
                    Log.Information("Showing settings window on startup (GCSM_SHOW_WINDOW_ON_START != 0).");
                    ShowForm();
                }
                else
                {
                    HideForm();
                }
            }
        }

        private void RunAtStartupCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            var regKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
            try
            {
                var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(regKey);
                if (runAtStartupCheckBox.Checked)
                {
                    // add to registry
                    regKeyAppRoot.SetValue("GoogleContactSync", "\"" + Application.ExecutablePath + "\"");
                }
                else
                {
                    // remove from registry
                    regKeyAppRoot.DeleteValue("GoogleContactSync");
                }
            }
            catch (Exception ex)
            {
                //if we can't write to that key, disable it... 
                runAtStartupCheckBox.Checked = false;
                runAtStartupCheckBox.Enabled = false;
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(new Exception("Error saving 'Run program at startup' settings into Registry key '" + regKey + "' Error: " + ex.Message, ex));
            }
        }

        private void ValidateSyncButton()
        {
            syncButton.Enabled = ValidUserName && ValidProfile && ValidSyncFolders;

            SetBgColor(contactFoldersComboBox, ValidSyncContactFolders);
            SetBgColor(btSyncContacts, ValidSyncContactFolders);
            SetBgColor(appointmentFoldersComboBox, ValidSyncOutlookAppointmentFolders);
            SetBgColor(appointmentGoogleFoldersComboBox, ValidSyncGoogleAppointmentFolders);
            SetBgColor(btSyncAppointments, ValidSyncAppointmentFolders);
        }

        private void Donate_Click(object sender, EventArgs e)
        {
            Process.Start("https://sourceforge.net/project/project_donations.php?group_id=369321");
        }

        private void Donate_MouseEnter(object sender, EventArgs e)
        {
            Donate.BackColor = Color.LightGray;
        }

        private void Donate_MouseLeave(object sender, EventArgs e)
        {
            Donate.BackColor = Color.Transparent;
        }

        private void HideButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ProxySettingsLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_proxy != null)
            {
                _proxy.ShowDialog(this);
            }
        }

        private void SettingsForm_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            ShowHelp();
        }

        private void SettingsForm_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            ShowHelp();
        }

        private static void ShowHelp()
        {
            // go to the page showing the help and howto instructions
            Process.Start("https://googlesyncmod.sourceforge.io/");
        }

        private void BtSyncContacts_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on appointments for syncing now).", "No sync switched on");
                btSyncAppointments.Checked = true;
            }
            contactFoldersComboBox.Visible = btSyncContacts.Checked;
            btSyncContactsForceRTF.Visible = btSyncContacts.Checked;
            SyncPhotosCheckBox.Visible = btSyncContacts.Checked;
            ValidateSyncButton();
        }

        private void BtSyncAppointments_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on contacts for syncing now).", "No sync switched on");
                btSyncContacts.Checked = true;
            }
            appointmentFoldersComboBox.Visible = appointmentGoogleFoldersComboBox.Visible = btSyncAppointments.Checked;
            labelTimezone.Visible = btMonthsPast.Visible = btMonthsFuture.Visible = btSyncAppointments.Checked;
            appointmentTimezonesComboBox.Visible = btSyncAppointments.Checked;
            pastMonthInterval.Visible = btSyncAppointments.Checked && btMonthsPast.Checked;
            futureMonthInterval.Visible = btSyncAppointments.Checked && btMonthsFuture.Checked;
            btSyncAppointmentsForceRTF.Visible = btSyncAppointments.Checked;
            btSyncAppointmentsPrivate.Visible = SyncRemindersCheckBox.Visible = btSyncAppointments.Checked;
            IncludePastRemindersCheckBox.Visible = btSyncAppointments.Checked && SyncRemindersCheckBox.Checked;
            ValidateSyncButton();
        }

        private void SyncReminders_CheckedChanged(object sender, EventArgs e)
        {
            IncludePastRemindersCheckBox.Visible = btSyncAppointments.Checked && SyncRemindersCheckBox.Checked;
        }

        private void CmbSyncProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            var comboBox = (ComboBox)sender;

            if ((0 == comboBox.SelectedIndex) || (comboBox.SelectedIndex == (comboBox.Items.Count - 1)))
            {
                using (var _configs = new ConfigurationManagerForm())
                {
                    if (0 == comboBox.SelectedIndex && _configs != null)
                    {
                        SyncProfile = ConfigurationManagerForm.AddProfile();
                        ClearSettings();
                    }

                    if (comboBox.SelectedIndex == (comboBox.Items.Count - 1) && _configs != null)
                    {
                        _configs.CurrentSyncProfile = SyncProfile;
                        _configs.Synchronizer = sync;
                        _configs.ShowDialog(this);
                    }
                }
                FillSyncProfileItems();

                comboBox.Text = SyncProfile;
                SaveSettings();
            }
            if (comboBox.SelectedIndex < 0)
            {
                MessageBox.Show("Please select Sync Profile.", "No sync switched on");
            }
            else
            {
                if (cmbSyncProfile_PreviouslySelectedIndex != -1)
                {
                    SaveSettings(comboBox.Items[cmbSyncProfile_PreviouslySelectedIndex].ToString());

                }
                LoadSettings(comboBox.Text);
                LoadSettingsFolders(comboBox.Text);
                SyncProfile = comboBox.Text;
                cmbSyncProfile_PreviouslySelectedIndex = comboBox.SelectedIndex;
            }

            ValidateSyncButton();
        }

        private void ContacFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            folderComboBox_SelectedIndexChanged(sender, ref syncContactsFolder, "Outlook Contacts");
        }

        private void AppointmentFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            folderComboBox_SelectedIndexChanged(sender, ref syncAppointmentsFolder, "Outlook Appointments");
        }

        private void AppointmentGoogleFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var message = "Select the Google Calendar you want to sync";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is GoogleCalendar calendar)
            {
                syncAppointmentsGoogleFolder = comboBox.SelectedValue.ToString();
                toolTip.SetToolTip(comboBox, message + ":\r\n" + calendar.DisplayName);
            }
            else
            {
                syncAppointmentsGoogleFolder = "";
                toolTip.SetToolTip(comboBox, message);
            }

            ValidateSyncButton();            
        }

        private void folderComboBox_SelectedIndexChanged(object sender, ref string folder, string content)
        {
            string message = $"Select the {content} folder you want to sync to";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is OutlookFolder)
            {
                folder = (sender as ComboBox).SelectedValue.ToString();
                toolTip.SetToolTip((sender as ComboBox), message + ":\r\n" + ((OutlookFolder)(sender as ComboBox).SelectedItem).DisplayName);
            }
            else
            {
                folder = "";
                toolTip.SetToolTip((sender as ComboBox), message);
            }

            ValidateSyncButton();
        }

        private void BtSyncDelete_CheckedChanged(object sender, EventArgs e)
        {
            btPromptDelete.Visible = btSyncDelete.Checked;
            btPromptDelete.Checked = btSyncDelete.Checked;
        }

        private void PictureBoxExit_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Do you really want to exit " + Application.ProductName + "? This will also stop the service performing automatic synchronizaton in the background. If you only want to hide the settings form, use the 'Hide' Button instead.", "Exit " + Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                CancelButton_Click(sender, EventArgs.Empty); //Close running thread
                requestClose = true;
                Close();
            }
        }

        private void SystemEvents_PowerModeSwitch(object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode == PowerModes.Suspend)
            {
                TimerSwitch(false);
            }
            else if (e.Mode == PowerModes.Resume)
            {
                TimerSwitch(true);
            }
        }

        private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                boolShowBalloonTip = false;
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                boolShowBalloonTip = true;
                TimerSwitch(true);
            }
        }

        private void AutoSyncInterval_ValueChanged(object sender, EventArgs e)
        {
            TimerSwitch(true);
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            cancellationTokenSource.Cancel();
            KillSyncThread();
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand, ControlThread = true)]
        private void KillSyncThread()
        {
            if (syncThread != null && syncThread.IsAlive)
            {
                syncThread.Abort();
            }
        }

        #region syncing icon
        public void IconTimerSwitch(bool value)
        {
            if (InvokeRequired)
            {
                var h = new SwitchHandler(IconTimerSwitch);
                Invoke(h, new object[] { value });
            }
            else
            {
                if (value) //Reset Icon to default icon as starting point for the syncing icon
                {
                    notifyIcon.Icon = Icon0;
                }

                iconTimer.Enabled = value;
            }
        }

        private void IconTimer_Tick(object sender, EventArgs e)
        {
            ShowNextIcon();
        }

        private void ShowNextIcon()
        {
            if (InvokeRequired)
            {
                var h = new IconHandler(ShowNextIcon);
                Invoke(h, new object[] { });
            }
            else
            {
                notifyIcon.Icon = GetNextIcon(notifyIcon.Icon);
            };
        }

        private Icon GetNextIcon(Icon currentIcon)
        {
#pragma warning disable IDE0046 // Convert to conditional expression
            if (currentIcon == IconError) //Don't change the icon anymore, once an error occurred
            {
                return IconError;
            }

            if (currentIcon == Icon30)
            {
                return Icon60;
            }

            if (currentIcon == Icon60)
            {
                return Icon90;
            }

            if (currentIcon == Icon90)
            {
                return Icon120;
            }

            if (currentIcon == Icon120)
            {
                return Icon150;
            }

            if (currentIcon == Icon150)
            {
                return Icon180;
            }

            if (currentIcon == Icon180)
            {
                return Icon210;
            }

            if (currentIcon == Icon210)
            {
                return Icon240;
            }

            if (currentIcon == Icon240)
            {
                return Icon270;
            }

            if (currentIcon == Icon270)
            {
                return Icon300;
            }

            if (currentIcon == Icon300)
            {
                return Icon330;
            }

            if (currentIcon == Icon330)
            {
                return Icon0;
            }

            return Icon30;
#pragma warning restore IDE0046 // Convert to conditional expression
        }
        #endregion

        private void AppointmentTimezonesComboBox_TextChanged(object sender, EventArgs e)
        {
            Timezone = appointmentTimezonesComboBox.Text;
        }

        private void LinkLabelRevokeAuthentication_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RevokeAuthentication();
        }

        public static void RevokeAuthentication()
        {
            try
            {
                Log.Information("Trying to remove Authentication...");
                var Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\GoContactSyncMOD\\";
                var AuthFolder = Folder + "\\Auth\\";
                var fDS = new FileDataStore(AuthFolder, true);
                fDS.ClearAsync();
                Log.Information("Removed Authentication...");
            }
            catch (Exception ex)
            {
                Log.Error($"Error revoking Google Authentication: {ex.Message}");
                Log.Debug(ex, "Exception");
            }
        }

        private void LoadAppointmentGoogleFoldersComboBox()
        {
            Log.Debug("Loading Google Appointments folders...");
            var googleAppointmentFolders = new ArrayList();

            appointmentGoogleFoldersComboBox.BeginUpdate();

            var defaultText = "    --- Select a Google Appointment folder ---";

            if (sync == null)
            {
                sync = new Synchronizer();
            }

            sync.SyncAppointments = btSyncAppointments.Checked;
            sync.LoginToGoogle(UserName.Text);

            if (sync.CalendarList != null)
            {
                foreach (var calendar in sync.CalendarList)
                {
                    googleAppointmentFolders.Add(new GoogleCalendar(calendar.Summary, calendar.Id, calendar.Primary ?? false));
                }
            }

            if (googleAppointmentFolders != null)
            {
                googleAppointmentFolders.Sort();
                googleAppointmentFolders.Insert(0, new GoogleCalendar(defaultText, defaultText, false));
                appointmentGoogleFoldersComboBox.DataSource = googleAppointmentFolders;
                appointmentGoogleFoldersComboBox.DisplayMember = "DisplayName";
                appointmentGoogleFoldersComboBox.ValueMember = "FolderID";
            }
            appointmentGoogleFoldersComboBox.EndUpdate();
            appointmentGoogleFoldersComboBox.SelectedValue = defaultText;

            //Select Default Folder per Default
            foreach (GoogleCalendar folder in appointmentGoogleFoldersComboBox.Items)
            {
                if (folder.IsDefaultFolder)
                {
                    appointmentGoogleFoldersComboBox.SelectedValue = folder.FolderID;
                    break;
                }
            }

            Log.Debug("Loaded Google Appointments folders");
        }

        private void AppointmentGoogleFoldersComboBox_Enter(object sender, EventArgs e)
        {
            if (appointmentGoogleFoldersComboBox.DataSource == null ||
                appointmentGoogleFoldersComboBox.Items.Count <= 1)
            {
                LoadAppointmentGoogleFoldersComboBox();
            }
        }

        private void AutoSyncInterval_Enter(object sender, EventArgs e)
        {
            syncTimer.Enabled = false;
        }

        private void AutoSyncInterval_Leave(object sender, EventArgs e)
        {
            syncTimer.Enabled = true;
        }

        private void NewVersionLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (((LinkLabel)sender).LinkColor == Color.Red)
            {
                Log.Debug("Process Start for https://sourceforge.net/projects/googlesyncmod/files/latest/download");
                Process.Start("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
            }
            else
            {
                Log.Debug("Process Start for https://sourceforge.net/projects/googlesyncmod/");
                Process.Start("https://sourceforge.net/projects/googlesyncmod/");
            }
        }

        private void UserName_Validating(object sender, CancelEventArgs e)
        {
            var isValid = ValidUserName;//Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);

            //SetBgColor(UserName, isValid);
            syncButton.Enabled = isValid;

            if (isValid)
            {
                //toolTip.SetToolTip(UserName, string.Empty);
            }
            else
            {
                //toolTip.SetToolTip(UserName, "User is of wrong format, should be full Google Mail address, e.g. user@gmail.com");
                Log.Warning("User is of wrong format, should be full Google Mail address, e.g. user@gmail.com");
                e.Cancel = true;
            }
        }

        private void UserName_Validated(object sender, EventArgs e)
        {
            var _profile = cmbSyncProfile.Text;
            if (!string.IsNullOrEmpty(_profile))
            {
                if (prevUserName != UserName.Text)
                {
                    prevUserName = UserName.Text;
                    LoadAppointmentGoogleFoldersComboBox();
                }
            }
        }

        private void btMonthsPast_CheckedChanged(object sender, EventArgs e)
        {
            pastMonthInterval.Visible = btMonthsPast.Checked;
            futureMonthInterval.Visible = btMonthsFuture.Checked;
        }

        private void btMonthsFuture_CheckedChanged(object sender, EventArgs e)
        {
            pastMonthInterval.Visible = btMonthsPast.Checked;
            futureMonthInterval.Visible = btMonthsFuture.Checked;
        }

        private void futureMonthInterval_ValueChanged(object sender, EventArgs e)
        {
            //if negative futureMonthInterval selected (i.e. sync not until present), the pastMonthInterval must be more in the past
            if (futureMonthInterval.Value < 0 && futureMonthInterval.Value < -pastMonthInterval.Value)
                pastMonthInterval.Value = -futureMonthInterval.Value;
        }
    }
}
