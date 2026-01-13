using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Requests;
using Google.Apis.Util.Store;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.PeopleService.v1;

using Microsoft.Office.Interop.Outlook;
using Polly;

using Polly.Retry;
using Serilog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Event = Google.Apis.Calendar.v3.Data.Event;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;
using Polly.Registry;
using Polly.Contrib.WaitAndRetry;
using static Google.Apis.PeopleService.v1.PeopleResource.ConnectionsResource;
using Polly.Wrap;

namespace GoContactSyncMod
{
    internal class Synchronizer : IDisposable
    {
        public const int OutlookUserPropertyMaxLength = 32;
        public const string OutlookUserPropertyPrefixTemplate = "g/con/";
        public const string OutlookUserPropertyTemplate = OutlookUserPropertyPrefixTemplate + "{0}/";
        //public const string myContactsGroup = "System ContactGroup: My Contacts";        
        internal const string myContactsGroup = "contactGroups/myContacts";
        //ToDo: Not used UpdatePersonFields: calendarUrls,externalIds,genders,interests,locales,miscKeywords,occupations,sipAddresses
        public const string UPDATE_PERSON_FIELDS = "addresses,biographies,birthdays,clientData,emailAddresses,events,imClients,locations,memberships,names,nicknames,organizations,phoneNumbers,relations,urls,userDefined,fileAses";
        //ToDo: Not used GetPersonFields: ageRanges,calendarUrls,coverPhotos,externalIds,genders,interests,locales,miscKeywords,sipAddresses,skills,occupations
        public const string GET_PERSON_FIELDS = UPDATE_PERSON_FIELDS + ",metadata,photos";
        private static readonly object _syncRoot = new object();
        internal static string UserName;

        private readonly PolicyRegistry registry = null;
        private readonly PolicyRegistry registryWrapPolicies = null;

        public int TotalCount { get; set; }
        public int SyncedCount { get; private set; }
        public int DeletedCount { get; private set; }
        public int ErrorCount { get; private set; }
        public int SkippedCount { get; set; }
        public int SkippedCountNotMatches { get; set; }
        public ConflictResolution ConflictResolution { get; set; }
        public DeleteResolution DeleteGoogleResolution { get; set; }
        public DeleteResolution DeleteOutlookResolution { get; set; }

        public delegate void NotificationHandler(string message);

        public delegate void DuplicatesFoundHandler(string title, string message);
        public delegate void ErrorNotificationHandler(string title, Exception ex);
        public delegate void TimeZoneNotificationHandler(string timeZone);

        public event DuplicatesFoundHandler DuplicatesFound;
        public event ErrorNotificationHandler ErrorEncountered;
        public event TimeZoneNotificationHandler TimeZoneChanges;

        public PeopleServiceService GooglePeopleService { get; private set; }
        public PeopleResource GooglePeopleResource { get; private set; }
        public ListRequest GooglePeopleRequest { get; private set; }

        private EventsResource GoogleEventsResource { get; set; }

        private static NameSpace _OutlookNameSpace;
        public static NameSpace OutlookNameSpace
        {
            get
            {
                //Just create outlook instance again, in case the namespace is null
                CreateOutlookInstance();
                return _OutlookNameSpace;
            }
        }

        public static Outlook.Application OutlookApplication { get; private set; }
        public Items OutlookContacts { get; private set; }
        public Items OutlookAppointments { get; private set; }
        public Collection<ContactMatch> OutlookContactDuplicates { get; set; }
        public Collection<ContactMatch> GoogleContactDuplicates { get; set; }

        public Collection<Person> GoogleContacts { get; private set; }
        private CalendarService GoogleCalendarService;
        public Collection<Event> GoogleAppointments { get; private set; }
        public Collection<Event> AllGoogleAppointments { get; private set; }
        public IList<CalendarListEntry> CalendarList { get; private set; }
        public Collection<ContactGroup> GoogleGroups { get; set; }
        public static string OutlookPropertyPrefix { get; private set; }

        public static string OutlookPropertyNameId => OutlookPropertyPrefix + "id";

        public static string OutlookPropertyNameSynced => OutlookPropertyPrefix + "up";
        public static string OutlookPropertyNameEtag => OutlookPropertyPrefix + "etag";
        //public static string OutlookPropertyNameSource => OutlookPropertyPrefix + "source";
        public SyncOption SyncOption { get; set; } = SyncOption.MergeOutlookWins;

        public static string SyncProfile { get; set; }
        public static string SyncContactsFolder { get; set; }
        public static string SyncAppointmentsFolder { get; set; }
        public static string SyncAppointmentsGoogleFolder { get; set; }
        public static string SyncAppointmentsGoogleTimeZone { get; set; }

        public static ushort MonthsInPast { get; set; }
        public static bool RestrictMonthsInPast { get; set; }
        public static short MonthsInFuture { get; set; }
        public static bool RestrictMonthsInFuture { get; set; }
        public static string Timezone { get; set; }
        public static bool MappingBetweenTimeZonesRequired { get; set; }

        public List<ContactMatch> Contacts { get; private set; }

        public List<AppointmentMatch> Appointments { get; private set; }

        private HashSet<string> ContactExtendedPropertiesToRemoveIfTooMany = null;
        private HashSet<string> ContactExtendedPropertiesToRemoveIfTooBig = null;
        private HashSet<string> ContactExtendedPropertiesToRemoveIfDuplicated = null;

        /// <summary>
        /// If true deletes contacts if synced before, but one is missing. Otherwise contacts will bever be automatically deleted
        /// </summary>
        public bool SyncDelete { get; set; }
        public bool PromptDelete { get; set; }

        /// <summary>
        /// If true sync also contacts
        /// </summary>
        public bool SyncContacts { get; set; }
        public static bool SyncContactsForceRTF { get; set; }
        public static bool SyncPhotos { get; set; }

        /// <summary>
        /// If true sync also appointments (calendar)
        /// </summary>
        public bool SyncAppointments { get; set; }
        public static bool SyncAppointmentsForceRTF { get; set; }
        public static bool SyncAppointmentsPrivate { get; set; }
        public static bool SyncReminders { get; set; }
        public static bool IncludePastReminders { get; set; }
        /// <summary>
        /// if true, use Outlook's FileAs for Google Title/FullName. If false, use Outlook's Fullname
        /// </summary>
        public bool UseFileAs { get; set; }

        public Synchronizer()
        {
            var delay = Backoff.ConstantBackoff(TimeSpan.FromMilliseconds(200), retryCount: 5, fastFirst: true);

            var policy = Policy
                .Handle<TaskCanceledException>()
                .WaitAndRetry(delay, onRetry: (exception, retryCount, context) =>
                {
                    Log.Debug("Retry");
                });

            registry = new PolicyRegistry()
            {
                { "Standard", policy },
                { "Contact Read", CreatGoogleContactReadRetryPolicies() }
            };

            registryWrapPolicies = new PolicyRegistry()
            {
                { "Contact Write", CreatGoogleContactWriteRetryPolicies() }
            };
        }

        private Policy CreatGoogleContactReadRetryPolicies()
        {
            var delay = Backoff.ExponentialBackoff(TimeSpan.FromMilliseconds(1000), retryCount: 5);

            var policyContactRead = Policy
                .Handle<Google.GoogleApiException>(ex => ex.HttpStatusCode == (HttpStatusCode)429 && ex.Error.Message.StartsWith("Quota exceeded for quota metric"))
                .WaitAndRetry(delay, onRetry: (exception, retryCount, context) =>
                {
                    Log.Debug($"Retry, waiting for {retryCount}");
                });

            return policyContactRead;
        }

        private PolicyWrap CreatGoogleContactWriteRetryPolicies()
        {
            var delay = Backoff.ExponentialBackoff(TimeSpan.FromMilliseconds(1000), retryCount: 5);

            var policyContactWrite = Policy
                .Handle<Google.GoogleApiException>(ex => ex.HttpStatusCode == (HttpStatusCode)429 && ex.Error.Message.StartsWith("Quota exceeded for quota metric"))
                .WaitAndRetry(delay, onRetry: (exception, retryCount, context) =>
                {
                    Log.Debug($"Retry, waiting for {retryCount}");
                });

            // http://stackoverflow.com/questions/23804960/contactsrequest-insertfeeduri-newentry-sometimes-fails-with-system-net-protoc
            var policyProtocolViolationException = Policy
                .Handle<ProtocolViolationException>()
                .Retry(1, onRetry: (exception, retryCount) =>
                {
                    Log.Debug($"Retry {retryCount}");
                });

            return Policy.Wrap(policyProtocolViolationException, policyContactWrite);
        }

        public void LoginToGoogle(string username)
        {
            Log.Information("Connecting to Google...");

            //check if it is now relogin to different user
            if (username != UserName)
            {
                GooglePeopleResource = null;
                GooglePeopleRequest = null;
                GooglePeopleService = null;
                GoogleEventsResource = null;
                SyncAppointmentsGoogleFolder = null;
            }

            if (((GooglePeopleResource == null || GooglePeopleRequest == null) && SyncContacts) || GoogleEventsResource == null && SyncAppointments)
            {
                //OAuth2 for all services
                var scopes = new List<string>
                {
                    CalendarService.Scope.Calendar,
                    PeopleServiceService.Scope.Contacts
                };

                //take user credentials
                UserCredential credential;

                var Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\GoContactSyncMOD\\";
                var AuthFolder = Folder + "Auth\\";

                Stream stream;

                /* TODO user wants/needs to use own client secrets - e.g. api limits reached 
                 * Not used at the moment
                 * 
                var ClientSecretsFile = Folder + "\\Client-Secrets\\client_secrets.json";

                if (File.Exists(ClientSecretsFile))
                {
                    stream = new FileStream(ClientSecretsFile, FileMode.Open);
                }
                else
                {
                    //load client secret from ressources
                    byte[] jsonSecrets = Properties.Resources.client_secrets;
                    stream = new MemoryStream(jsonSecrets);
                }
                */

                //load client secret provided by gcsm from ressources
                var jsonSecrets = Properties.Resources.client_secrets;
                stream = new MemoryStream(jsonSecrets);

                //cancel auth request after 60 seconds
                var authTimeout = 60;

                try
                {
                    var fDS = new FileDataStore(AuthFolder, true);

                    var clientSecrets = GoogleClientSecrets.FromStream(stream);

                    using (var cts = new CancellationTokenSource())
                    {
                        //Cancel auth request after timeout
                        cts.CancelAfter(TimeSpan.FromSeconds(authTimeout));
                        var ct = cts.Token;
                        ct.ThrowIfCancellationRequested();

                        credential = GCSMOAuth2WebAuthorizationBroker.AuthorizeAsync(
                                clientSecrets.Secrets,
                                scopes.ToArray(),
                                username,
                                ct,
                                fDS).
                                Result;

                        var initializer = new Google.Apis.Services.BaseClientService.Initializer
                        {
                            HttpClientInitializer = credential
                        };

                        //var parameters = new OAuth2Parameters  //ToDo: Check, if still needed for new Google Api
                        //{
                        //    ClientId = clientSecrets.Secrets.ClientId,
                        //    ClientSecret = clientSecrets.Secrets.ClientSecret,

                        //    // Note: AccessToken is valid only for 60 minutes
                        //    AccessToken = credential.Token.AccessToken,
                        //    RefreshToken = credential.Token.RefreshToken
                        //};
                        //Log.Information(Application.ProductName);
                        //var settings = new RequestSettings(
                        //    Application.ProductName, parameters);

                        if (SyncContacts)
                        {
                            //PeopleRequest = new PeopleRequest(rs);
                            //PeopleRequest = new PeopleRequest(settings);
                            GooglePeopleService = GoogleServices.CreatePeopleService(initializer);
                            GooglePeopleResource = new PeopleResource(GooglePeopleService);
                            GooglePeopleRequest = GooglePeopleResource.Connections.List("people/me");
                            GooglePeopleRequest.PersonFields = GET_PERSON_FIELDS;
                        }

                        if (SyncAppointments)
                        {
                            GoogleCalendarService = GoogleServices.CreateCalendarService(initializer);

                            const int NumberOfRetries = 3;
                            const int DelayOnRetry = 1000;
                            for (var i = 1; i <= NumberOfRetries; ++i)
                            {
                                try
                                {
                                    CalendarList = GoogleCalendarService.CalendarList.List().Execute().Items;
                                    break;
                                }
                                catch (Exception ex) when (i < NumberOfRetries)
                                {
                                    Log.Debug(ex, $"Try {i}");
                                    Task.Delay(DelayOnRetry);
                                }
                            }

                            //Get Primary Calendar, if not set from outside
                            if (string.IsNullOrEmpty(SyncAppointmentsGoogleFolder))
                            {
                                foreach (var calendar in CalendarList)
                                {
                                    if (calendar.Primary != null && calendar.Primary.Value)
                                    {
                                        SyncAppointmentsGoogleFolder = calendar.Id;
                                        SyncAppointmentsGoogleTimeZone = calendar.TimeZone;
                                        if (string.IsNullOrEmpty(SyncAppointmentsGoogleTimeZone))
                                        {
                                            Log.Debug($"Empty Google time zone for calendar {calendar.Id}");
                                        }

                                        break;
                                    }
                                }
                            }
                            else
                            {
                                var found = false;
                                foreach (var calendar in CalendarList)
                                {
                                    if (calendar.Id == SyncAppointmentsGoogleFolder)
                                    {
                                        SyncAppointmentsGoogleTimeZone = calendar.TimeZone;
                                        if (string.IsNullOrEmpty(SyncAppointmentsGoogleTimeZone))
                                        {
                                            Log.Debug($"Empty Google time zone for calendar {calendar.Id}");
                                        }
                                        else
                                        {
                                            found = true;
                                        }

                                        break;
                                    }
                                }
                                if (!found)
                                {
                                    Log.Warning($"Cannot find calendar, id is {SyncAppointmentsGoogleFolder}");

                                    Log.Debug("Listing calendars:");
                                    foreach (var calendar in CalendarList)
                                    {
                                        if (calendar.Primary != null && calendar.Primary.Value)
                                        {
                                            Log.Debug($"Id (primary): {calendar.Id}");
                                        }
                                        else
                                        {
                                            Log.Debug($"Id: {calendar.Id}");
                                        }
                                    }
                                }
                            }

                            if (SyncAppointmentsGoogleFolder == null)
                            {
                                throw new Exception("Google Calendar not defined (primary not found)");
                            }

                            GoogleEventsResource = GoogleCalendarService.Events;
                        }
                    }
                }
                catch (Exception ex) when (ex.InnerException is OperationCanceledException)
                {
                    Log.Error($"Authorisation to allow GCSM to manage your Google calendar was cancelled. Hint: You have to answer the google consent screen within {authTimeout} seconds. {ex.InnerException.Message}");
                }
                catch (Exception ex)
                {
                    Log.Error($"Error logging into Google: {ex.Message}");
                    Log.Debug(ex, "Exception");

                }
                finally
                {
                    stream.Dispose();
                }
            }

            UserName = username;

            var maxUserIdLength = OutlookUserPropertyMaxLength - (OutlookUserPropertyTemplate.Length - 3 + 2);//-3 = to remove {0}, +2 = to add length for "id" or "up"
            var userId = username;
            if (userId.Length > maxUserIdLength)
            {
                userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.
            }
            //Remove characters not allowed for Outlook user property names: []_#
            userId = userId.Replace("#", "").Replace("[", "").Replace("]", "").Replace("_", "");

            OutlookPropertyPrefix = string.Format(OutlookUserPropertyTemplate, userId);
        }

        public void LoginToOutlook()
        {
            Log.Information("Connecting to Outlook...");

            try
            {
                CreateOutlookInstance();
            }
            catch (Exception e) when ((e is COMException) || (e is InvalidCastException))
            {
                try
                {
                    // If outlook was closed/terminated inbetween, we will receive an Exception
                    // System.Runtime.InteropServices.COMException (0x800706BA): The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)
                    // so recreate outlook instance
                    // And sometimes we receive an Exception
                    // System.InvalidCastException 0x8001010E (RPC_E_WRONG_THREAD))
                    Log.Information("Cannot connect to Outlook, creating new instance....");
                    Log.Debug(e, "Exception");

                    OutlookApplication = null;
                    _OutlookNameSpace = null;
                    CreateOutlookInstance();
                }
                catch (Exception ex)
                {
                    var message = $"Cannot connect to Outlook.\r\nPlease restart {Application.ProductName} and try again. If error persists, please inform developers on SourceForge.";
                    // Error again? We need full stacktrace, display it!
                    throw new Exception(message, ex);
                }
            }
        }

        private static void GetAlreadyApplicationIfStarted(int num_tries)
        {
            //Try to create new Outlook application few times, because mostly it fails the first time, if not yet running
            // First try to get the running application in case Outlook is already started
            for (var i = 0; i < num_tries; i++)
            {
                try
                {
                    OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                    if (OutlookApplication != null)
                    {
                        return;  //Exit, if creating outlook application was successful
                    }
                    else
                    {
                        OutlookApplication = new Outlook.Application();
                        if (OutlookApplication != null)
                        {
                            return;  //Exit, if creating outlook application was successful
                        }
                        Log.Debug($"CreateOutlookApplication (null), try: {i + 1}");
                        Thread.Sleep(1000 * 10 * (i + 1));
                    }
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x80029c4a)
                {
                    Log.Debug(ex, "CreateOutlookApplication (0x80029c4a)");
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x800401E3)
                {
                    var processes = Process.GetProcessesByName("OUTLOOK");
                    var p_num = processes.Length;

                    Log.Debug($"CreateOutlookApplication (0x800401E3), number of started processes: {p_num}");

                    if (p_num > 0)
                    {
                        try
                        {
                            Log.Debug($"CreateOutlookApplication (0x800401E3), number of started processes: {p_num}, 1st try");
                            OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                            if (OutlookApplication != null)
                            {
                                return;  //Exit, if creating outlook application was successful
                            }
                        }
                        catch (Exception ex1)
                        {
                            Log.Debug(ex1, $"CreateOutlookApplication (0x800401E3), number of started processes: {p_num}, 1st exception");
                        }

                        Thread.Sleep(1000 * 10);
                        try
                        {
                            Log.Debug($"CreateOutlookApplication (0x800401E3), number of started processes: {p_num}, 2nd try");
                            OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                            if (OutlookApplication != null)
                            {
                                return;  //Exit, if creating outlook application was successful
                            }
                        }
                        catch (Exception ex1)
                        {
                            Log.Debug(ex1, $"CreateOutlookApplication (0x800401E3), number of started processes: {p_num}, 2nd exception");
                        }
                    }

                    try
                    {
                        OutlookApplication = new Outlook.Application();
                    }
                    catch (COMException e) when ((uint)e.ErrorCode == 0x80080005)
                    {
                        Log.Debug("CreateOutlookApplication (0x80080005)");
                        throw new NotSupportedException("Outlook and \"" + Application.ProductName + "\" are started by different users. For example you run Outlook with the \"Run as administrator\" option and \"" + Application.ProductName + "\" as regular user (or the other way around). This is not supported.", e);
                    }
                    catch (Exception e)
                    {
                        Log.Debug(e, "Exception");
                    }

                    if (OutlookApplication != null)
                    {
                        return;  //Exit, if creating outlook application was successful
                    }
                    Thread.Sleep(1000 * 10 * (i + 1));
                }
                catch (COMException ex)
                {
                    Log.Debug(ex, "CreateOutlookApplication (COMException)");

                    try
                    {
                        OutlookApplication = new Outlook.Application();
                    }
                    catch (COMException e) when ((uint)e.ErrorCode == 0x80080005)
                    {

                        Log.Debug("CreateOutlookApplication (0x80080005)");
                        throw new NotSupportedException("Outlook and \"GO Contact Sync Mod\" are started by different users. For example you run Outlook with the \"Run as administrator\" option and \"GO Contact Sync Mod\" as regular user (or the other way around). This is not supported.", e);

                    }
                    catch (Exception e)
                    {
                        Log.Debug(e, "Exception");
                    }

                    if (OutlookApplication != null)
                    {
                        return;  //Exit, if creating outlook application was successful
                    }
                    Thread.Sleep(1000 * 10 * (i + 1));
                }
                catch (NotSupportedException)
                {
                    throw;
                }
                catch (InvalidCastException ex)
                {
                    Log.Debug(ex, "CreateOutlookApplication (InvalidCastException)");
                    throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                }
                catch (Exception ex) when (i == (num_tries - 1))
                {
                    Log.Debug(ex, "CreateOutlookApplication (Exception): last try");
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                }
                catch (Exception)
                {
                    Log.Debug($"CreateOutlookApplication (Exception), try: {i + 1}");
                    Thread.Sleep(1000 * 10 * (i + 1));
                }
            }
        }

        private static void CreateApplicationIfNotStarted(int num_tries)
        {
            // Next try to have new running instance of Outlook
            Log.Debug("CreateOutlookApplication: new Outlook.Application");
            for (var i = 0; i < num_tries; i++)
            {
                try
                {
                    OutlookApplication = new Outlook.Application();
                    if (OutlookApplication != null)
                    {
                        return;  //Exit, if creating outlook application was successful
                    }
                    else
                    {
                        Log.Debug($"CreateOutlookApplication (null), try: {i + 1}");
                        Thread.Sleep(1000 * 10 * (i + 1));
                    }
                }
                catch (COMException ex)
                {
                    if ((uint)ex.ErrorCode == 0x80029c4a)
                    {
                        Log.Debug(ex, "CreateOutlookApplication (0x80029c4a)");
                        throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                    }
                    Thread.Sleep(1000 * 10 * (i + 1));
                }
                catch (InvalidCastException ex)
                {
                    Log.Debug(ex, "CreateOutlookApplication (InvalidCastException)");
                    throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                }
                catch (Exception ex)
                {
                    if (i == (num_tries - 1))
                    {
                        Log.Debug(ex, "CreateOutlookApplication (Exception): last try");
                        throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                    }
                    else
                    {
                        Log.Debug($"CreateOutlookApplication (Exception), try: {i + 1}");
                        Thread.Sleep(1000 * 10 * (i + 1));
                    }
                }
            }
        }

        private static void CreateOutlookApplication()
        {
            const int num_tries = 3;

            GetAlreadyApplicationIfStarted(num_tries);
            if (OutlookApplication != null)
            {
                return;  //Exit, if creating outlook application was successful
            }
            CreateApplicationIfNotStarted(num_tries);
        }

        private static void CreateOutlookNamespace()
        {
            const int num_tries = 5;
            //Try to create new Outlook namespace few times, because mostly it fails the first time, if not yet running
            for (var i = 0; i < num_tries; i++)
            {
                try
                {
                    _OutlookNameSpace = OutlookApplication.GetNamespace("MAPI");
                    if (_OutlookNameSpace != null)
                    {
                        break;  //Exit the for loop, if getting outlook namespace was successful
                    }
                    else
                    {
                        Log.Debug($"CreateOutlookNamespace (null), try: {i + 1}");
                    }
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x80029c4a)
                {
                    Log.Debug(ex, "CreateOutlookNamespace (0x80029c4a)");
                    throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                }
                catch (COMException ex) when (i == (num_tries - 1))
                {
                    Log.Debug(ex, "CreateOutlookNamespace (COMException): last try");
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                }
                catch (COMException)
                {
                    Log.Debug($"CreateOutlookNamespace (COMException), try: {i + 1}");
                    Thread.Sleep(1000 * 10 * (i + 1));
                }
                catch (InvalidCastException ex)
                {
                    Log.Debug(ex, "CreateOutlookNamespace (InvalidCastException)");
                    throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                }
                catch (Exception ex) when (i == (num_tries - 1))
                {
                    Log.Debug(ex, "CreateOutlookNamespace (Exception): last try");
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);

                }
                catch (Exception)
                {
                    Log.Debug($"CreateOutlookNamespace (Exception), try: {i + 1}");
                    Thread.Sleep(1000 * 10 * (i + 1));
                }
            }
        }

        private static void CreateOutlookInstanceHelper()
        {
            if (OutlookApplication == null)
            {
                CreateOutlookApplication();
                if (OutlookApplication == null)
                {
                    throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");
                }
            }

            if (_OutlookNameSpace == null)
            {
                CreateOutlookNamespace();
                if (_OutlookNameSpace == null)
                {
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");
                }

                Log.Debug($"Connected to Outlook: {VersionInformation.GetOutlookVersion(OutlookApplication)}");

                // OutlookNameSpace.Accounts was introduced in later version of Outlook
                // calling this in older version (like Outlook 2003) will result in "Attempted to read or write protected memory"
                try
                {
                    if (_OutlookNameSpace.Accounts != null && _OutlookNameSpace.Accounts.Count > 1)
                    {
                        Log.Debug($"Multiple outlook accounts: {_OutlookNameSpace.Accounts.Count}");
                    }
                }
                catch (AccessViolationException)
                {
                }
            }
        }

        private static void CreateOutlookInstance()
        {
            CreateOutlookInstanceHelper();

            var retryCount = 0;
            while (retryCount < 10)
            {
                try
                {
                    if (string.IsNullOrEmpty(SyncContactsFolder))
                    {
                        _OutlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                    }
                    else
                    {
                        _OutlookNameSpace.GetFolderFromID(SyncContactsFolder);
                    }
                    return;
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x80040201)
                {
                    retryCount++;
                    Log.Debug("0x80040201 - LogoffOutlookNameSpace");
                    LogoffOutlookNameSpace();
                    Log.Debug("0x80040201 - CreateOutlookInstanceHelper");
                    CreateOutlookInstanceHelper();
                    Log.Debug("0x80040201 - GetFolder");
                    if (string.IsNullOrEmpty(SyncContactsFolder))
                    {
                        _OutlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                    }
                    else
                    {
                        _OutlookNameSpace.GetFolderFromID(SyncContactsFolder);
                    }
                    Log.Debug("0x80040201 - Done");
                    return;
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x80010001)
                {
                    retryCount++;
                    // RPC_E_CALL_REJECTED - sleep and retry
                    Log.Debug($"RPC_E_CALL_REJECTED, trying {retryCount}");
                    Thread.Sleep(1000);
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x80029c4a)
                {
                    Log.Debug(ex, "Exception");
                    throw new NotSupportedException(OutlookRegistryUtils.GetPossibleErrorDiagnosis(), ex);
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x80040111)
                {
                    try
                    {
                        Log.Debug("Trying to logon, 1st try");
                        _OutlookNameSpace.Logon("", "", false, false);
                        Log.Debug("1st try OK");
                    }
                    catch (Exception e1)
                    {
                        Log.Debug(e1, "Exception");
                        try
                        {
                            Log.Debug("Trying to logon, 2nd try");
                            _OutlookNameSpace.Logon("", "", true, true);
                            Log.Debug("2nd try OK");
                        }
                        catch (Exception e2)
                        {
                            Log.Debug(e2, "Exception");
                            throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", e2);
                        }
                    }
                }
                catch (COMException ex)
                {
                    Log.Debug(ex, "Exception");
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                }
            }
        }

        private static void LogoffOutlookNameSpace()
        {
            try
            {
                Log.Debug("Disconnecting from Outlook...");
                if (_OutlookNameSpace != null)
                {
                    _OutlookNameSpace.Logoff();
                }
            }
            catch (Exception)
            {
                // if outlook was closed inbetween, we get an System.InvalidCastException or similar exception, that indicates that outlook cannot be acced anymore
                // so as outlook is closed anyways, we just ignore the exception here
            }

            try
            {
                Log.Debug($"Total allocated memory before collection: {GC.GetTotalMemory(false):N0}");

                _OutlookNameSpace = null;
                OutlookApplication = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                Log.Debug($"Total allocated memory after collection: {GC.GetTotalMemory(false):N0}");
            }
            finally
            {
                Log.Debug("Disconnected from Outlook");
            }
        }

        public void LogoffOutlook()
        {
            try
            {
                Log.Debug("Disconnecting from Outlook...");
                if (_OutlookNameSpace != null)
                {
                    _OutlookNameSpace.Logoff();
                }
            }
            catch (Exception)
            {
                // if outlook was closed inbetween, we get an System.InvalidCastException or similar exception, that indicates that outlook cannot be acced anymore
                // so as outlook is closed anyways, we just ignore the exception here
            }

            try
            {
                Log.Debug($"Total allocated memory before collection: {GC.GetTotalMemory(false):N0}");

                OutlookContactDuplicates = null;
                GoogleContactDuplicates = null;
                GoogleContacts = null;
                GoogleCalendarService = null;
                GooglePeopleService = null;
                GoogleAppointments = null;
                AllGoogleAppointments = null;
                CalendarList = null;
                GoogleGroups = null;
                Contacts = null;
                Appointments = null;
                ContactExtendedPropertiesToRemoveIfTooMany = null;
                ContactExtendedPropertiesToRemoveIfTooBig = null;
                ContactExtendedPropertiesToRemoveIfDuplicated = null;
                OutlookContacts = null;
                OutlookAppointments = null;
                _OutlookNameSpace = null;
                OutlookApplication = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                Log.Debug($"Total allocated memory after collection: {GC.GetTotalMemory(false):N0}");
            }
            finally
            {
                Log.Debug("Disconnected from Outlook");
            }
        }

        public void LogoffGoogle()
        {
            GooglePeopleResource = null;
            GooglePeopleRequest = null;
            GoogleEventsResource = null;
        }

        private void LoadOutlookContacts()
        {
            Log.Information("Loading Outlook contacts...");
            OutlookContacts = GetContactItems();
            Log.Debug($"Outlook Contacts Found: {OutlookContacts.Count}");
        }

        private void LoadOutlookAppointments()
        {
            Log.Information("Loading Outlook appointments...");
            OutlookAppointments = GetAppointmentItems();
            Log.Debug($"Outlook Appointments Found: {OutlookAppointments.Count}");
        }

        public MAPIFolder GetAppoimentsFolder()
        {
            return GetMAPIFolder(OlDefaultFolders.olFolderCalendar, SyncAppointmentsFolder);
        }

        private MAPIFolder GetMAPIFolder(OlDefaultFolders outlookDefaultFolder, string syncFolder)
        {
            MAPIFolder mapiFolder;
            if (string.IsNullOrEmpty(syncFolder))
            {
                mapiFolder = OutlookNameSpace.GetDefaultFolder(outlookDefaultFolder);
                if (mapiFolder == null)
                {
                    throw new Exception($"Error getting Default OutlookFolder: {outlookDefaultFolder}");
                }
            }
            else
            {
                try
                {
                    mapiFolder = OutlookNameSpace.GetFolderFromID(syncFolder);
                    if (mapiFolder == null)
                    {
                        throw new Exception($"Error getting OutlookFolder: {syncFolder}");
                    }
                }
                catch (COMException ex)
                {
                    Log.Debug(ex, "Exception");
                    LogoffOutlook();
                    LoginToOutlook();
                    mapiFolder = OutlookNameSpace.GetFolderFromID(syncFolder);
                    if (mapiFolder == null)
                    {
                        throw new Exception($"Error getting OutlookFolder: {syncFolder}");
                    }
                }
            }

            return mapiFolder;
        }

        private Items GetAppointmentItems()
        {
            return GetOutlookItems(OlDefaultFolders.olFolderCalendar, SyncAppointmentsFolder);
        }
        private Items GetContactItems()
        {
            return GetOutlookItems(OlDefaultFolders.olFolderContacts, SyncContactsFolder);
        }

        private Items GetOutlookItems(OlDefaultFolders outlookDefaultFolder, string syncFolder)
        {
            var mapiFolder = GetMAPIFolder(outlookDefaultFolder, syncFolder);
            var items = mapiFolder.Items;
            if (items == null)
            {
                throw new Exception($"Error getting Outlook items from Outlook folder: {mapiFolder.Name}");
            }
            else
            {
                return items;
            }
        }

        private void ScanForInvalidContact()
        {
            var invalid_contact = string.Empty;
            long i = 1;

            Log.Debug("Checking started");

            try
            {
                //var GooglePeopleRequest = GooglePeopleResource.Connections.List("people/me");
                GooglePeopleRequest.PersonFields = new List<string>() { "names" };
                GooglePeopleRequest.PageToken = null;

                do
                {
                    var response = GooglePeopleRequest.Execute();

                    if (response != null && response.Connections != null && response.Connections.Count > 0)
                    {
                        foreach (var person in response.Connections)
                        {
                            var name = ContactPropertiesUtils.GetGooglePrimaryName(person);
                            if (name != null)
                            {
                                invalid_contact = name.DisplayName;
                                invalid_contact = invalid_contact.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                                if (!string.IsNullOrWhiteSpace(invalid_contact))
                                {
                                    Log.Debug($"Checking ({i}): {invalid_contact}");
                                }
                                else
                                {
                                    Log.Debug($"Checking ({i}): N/A");
                                }
                                i++;

                                //if (name.Metadata != null)
                                //{
                                //    if (name.Metadata.Source != null)
                                //    {
                                //        if (name.Metadata.Source.Id is string id)
                                //        {
                                //            if (!string.IsNullOrWhiteSpace(id))
                                //            {
                                //                //var uri = new Uri(ContactsQuery.CreateContactsUri("default") + "/" + id);
                                //                var contact = GooglePeopleResource.Get(@"people/"+id);
                                var request = GooglePeopleResource.Get(person.ResourceName); //ToDo: Check
                                request.PersonFields = Synchronizer.GET_PERSON_FIELDS;
                                var contact = request.Execute();
                                Thread.Sleep(2000);
                                //            }
                                //        }
                                //    }
                                //}
                            }
                        }

                        GooglePeopleRequest.PageToken = response.NextPageToken;
                    }

                } while (!string.IsNullOrEmpty(GooglePeopleRequest.PageToken));

                Log.Debug("Checking finished");

            }
            catch (Google.GoogleApiException ex) //ToDo: Check counterpart of ClientFeedException in Google People Api (is it really GoogleApiException?)
            {
                if (ex.InnerException is FormatException)
                {
                    if (!string.IsNullOrWhiteSpace(invalid_contact))
                    {
                        Log.Error($"Error parsing contact: {invalid_contact}. Please check if contact has some date fields (like birthday) with ill formed date.");
                        return;
                    }
                    else
                    {
                        Log.Error("Error parsing contact: N/A. Please check if contact has some date fields (like birthday) with ill formed date.");
                        return;
                    }
                }
                Log.Debug(ex, "Exception");
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
            }
        }

        internal Person LoadGoogleContacts(string id)
        {
            const string message = "Error Loading Google Contacts. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
            const string service = "GCSM.Synchronizer.LoadGoogleContacts";

            Person ret = null;
            var googleContacts = new Collection<Person>();
            try
            {
                if (id == null) // Only log, if not specific Google Contacts are searched                    
                {
                    Log.Information("Loading Google Contacts...");
                }
                else
                {
                    Log.Debug($"Loading Google Contact with id {id}...");
                }

                //var group = GetGoogleGroupByName(myContactsGroup);
                const int num_tries = 5;
                var policy = registry.Get<Policy>("Contact Read");


                for (var i = 0; i < num_tries; i++)
                {
                    try
                    {
                        GooglePeopleRequest.PageToken = null;
                        GooglePeopleRequest.PageSize = 256;

                        do
                        {
                            ListConnectionsResponse response = null;
                            var result = policy.ExecuteAndCapture(() =>
                            {
                                response = GooglePeopleRequest.Execute();
                            });

                            if (response?.Connections != null)
                            {
                                foreach (var a in response.Connections)
                                {
                                    if (a.Metadata != null && !(a.Metadata.Deleted ?? false) && !googleContacts.Contains(a))
                                    {
                                        googleContacts.Add(a);
                                        if (!string.IsNullOrEmpty(id) && id.Equals(ContactPropertiesUtils.GetGoogleId(a), StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            ret = a;
                                            if (GoogleContacts == null)
                                                GoogleContacts = new Collection<Person>();
                                            if (!GoogleContacts.Contains(a))
                                                GoogleContacts.Add(a);//Only add found item to global GoogleContacts, if a single contact was searched for (e.g. not found before)

                                            Log.Debug($"Loaded Google Contact with id {id}.");

                                            return ret; //No need to further query the contacts, because the one with the passed id found

                                        }
                                    }
                                }
                            }
                            GooglePeopleRequest.PageToken = response.NextPageToken;
                        } while (!string.IsNullOrEmpty(GooglePeopleRequest.PageToken));

                        if (string.IsNullOrEmpty(id)) //Only update global GoogleContacts, if not a single contact was searched for
                        {
                            Log.Debug("Loaded Google Contacts.");
                            GoogleContacts = googleContacts;
                            return ret;
                        }
                        else //Contact not found => try again
                        {
                            Thread.Sleep(2000); //wait 2 seconds
                            Log.Debug($"LoadGoogleContacts, retry {i + 1} to find again the same contact with id {id}, because sometimes the first time not found for some magic reason...");
                        }

                    }
                    catch (ThreadAbortException)
                    {
                        Log.Debug($"LoadGoogleContacts, retry {i + 1}...");
                    }
                }
            }
            catch (Google.GoogleApiException ex) //ToDo: Check counterpart of ClientFeedException in Google People Api, Really GoogleApiException?
            {
                if (ex.InnerException is FormatException)
                {
                    Log.Error("One of your contacts at Google probably has invalid date inside one of date fields (for example birthday)");
                    ScanForInvalidContact();
                }
                //if (string.IsNullOrEmpty(id)) //Only update global GoogleContacts, if not a single contact was searched for
                //    GoogleContacts = googleContacts;
                throw;
            }
            catch (WebException ex)
            {
                //if (string.IsNullOrEmpty(id)) //Only update global GoogleContacts, if not a single contact was searched for
                //    GoogleContacts = googleContacts;
                throw new Google.GoogleApiException(service, message, ex); //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }
            catch (NullReferenceException ex)
            {
                //if (string.IsNullOrEmpty(id)) //Only update global GoogleContacts, if not a single contact was searched for
                //    GoogleContacts = googleContacts;
                throw new Google.GoogleApiException(service, message, new WebException("Error accessing feed", ex)); //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }

            //if (string.IsNullOrEmpty(id)) //Only update global GoogleContacts, if not a single contact was searched for
            //{
            //    GoogleContacts = googleContacts;
            //    Log.Debug("Loaded Google Contacts.");
            //}
            Log.Debug($"Google Contact with id {id} not found.");
            return ret;
        }


        public void LoadGoogleGroups()
        {
            var message = "Error Loading Google Groups. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
            var service = "GCSM.Synchronizer.LoadGoogleGroups";
            try
            {
                Log.Information("Loading Google Groups...");
                var groupsResource = new ContactGroupsResource(GooglePeopleService);

                //var query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"))
                //{
                //    NumberToRetrieve = 256,
                //    StartIndex = 0
                //};
                //query.ShowDeleted = false;

                GoogleGroups = new Collection<ContactGroup>();

                var groupsRequest = groupsResource.List();
                groupsRequest.PageToken = null;

                do
                {
                    var response = groupsRequest.Execute();

                    foreach (var a in response.ContactGroups)
                    {
                        GoogleGroups.Add(a);
                    }

                    groupsRequest.PageToken = response.NextPageToken;
                } while (!string.IsNullOrEmpty(groupsRequest.PageToken));
            }
            catch (WebException ex)
            {
                //Log.Error(message);
                throw new Google.GoogleApiException(service, message, ex);  //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }
            catch (NullReferenceException ex)
            {
                //Log.Error(message);
                throw new Google.GoogleApiException(service, message, new WebException("Error accessing Google Groups request", ex)); //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }
        }

        private void LoadGoogleAppointments()
        {
            Log.Information("Loading Google appointments...");

            const string message = "Error Loading Google appointments. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
            const string service = "GCSM.Synchronizer.LoadGoogleAppointments";

            try
            {

                GoogleAppointments = new Collection<Event>();

                var request = GoogleEventsResource.List(SyncAppointmentsGoogleFolder);

                request.PageToken = null;
                request.MaxResults = 256;

                //Only Load events from month range
                if (RestrictMonthsInPast)
                {
                    request.TimeMinDateTimeOffset = DateTime.Now.AddMonths(-MonthsInPast);
                }                

                if (RestrictMonthsInFuture)
                {
                    request.TimeMaxDateTimeOffset = DateTime.Now.AddMonths(MonthsInFuture);
                }
                

                Events response;

                do
                {

                    response = request.Execute();
                    foreach (var a in response.Items)
                    {
                        if ((a.RecurringEventId != null || !a.Status.Equals("cancelled")) &&
                            //a.Start != null && a.End != null &&
                            !GoogleAppointments.Contains(a) //ToDo: For an unknown reason, some appointments are duplicate in GoogleAppointments, therefore remove all duplicates before continuing  
                            )
                        {//only return not yet cancelled events (except for recurrence exceptions) and events not already in the list
                            GoogleAppointments.Add(a);                            
                        }
                        //else
                        //{
                        //    Log.Information("Skipped Appointment because it was cancelled on Google side: " + a.Summary + " - " + GetTime(a));
                        //SkippedCount++;
                        //}
                    }
                    request.PageToken = response.NextPageToken;
                }
                while (!string.IsNullOrEmpty(request.PageToken));
            }
            catch (WebException ex)
            {
                //Log.Error(message);
                throw new Google.GoogleApiException(service, message, ex);  //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }
            catch (NullReferenceException ex)
            {
                //Log.Error(message);
                throw new Google.GoogleApiException(service, message, new WebException("Error accessing feed", ex));    //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }

            //Remember, if all Google Appointments have been loaded
            if (!RestrictMonthsInPast && !RestrictMonthsInFuture)
            {
                AllGoogleAppointments = GoogleAppointments;
            }

            Log.Debug("Google Appointments Found: " + GoogleAppointments.Count);
        }

        /// <summary>
        /// Resets Google appointment matches.
        /// </summary>
        /// <param name="deleteGoogleAppointments">Should Google appointments be updated or deleted.</param>        
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A task that represents the asynchronous operation.</returns>
        internal async Task ResetGoogleAppointmentMatches(bool deleteGoogleAppointments, CancellationToken cancellationToken)
        {
            const int num_retries = 5;
            Log.Information("Processing Google appointments.");

            AllGoogleAppointments = null;
            GoogleAppointments = null;

            // First run batch updates, but since individual requests are not retried in case of any error rerun 
            // updates in single mode
            if (await BatchResetGoogleAppointmentMatches(deleteGoogleAppointments, cancellationToken))
            {
                // in case of error retry single updates five times
                for (var i = 1; i < num_retries; i++)
                {
                    if (!await SingleResetGoogleAppointmentMatches(deleteGoogleAppointments, cancellationToken))
                    {
                        break;
                    }
                }
            }

            Log.Information("Finished all Google changes.");
        }


        /// <summary>
        /// Resets Google appointment matches via single updates.
        /// </summary>
        /// <param name="deleteGoogleAppointments">Should Google appointments be updated or deleted.</param>        
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>If error occured.</returns>
        internal async Task<bool> SingleResetGoogleAppointmentMatches(bool deleteGoogleAppointments, CancellationToken cancellationToken)
        {
            const string message = "Error resetting Google appointments.";
            const string service = "GCSM.Synchronizer.SingleResetGoogleAppointmentMatches";

            var key = OutlookPropertiesUtils.GetKey();

            try
            {
                var request = GoogleEventsResource.List(SyncAppointmentsGoogleFolder);
                request.PageToken = null;

                if (RestrictMonthsInPast)
                {
                    request.TimeMinDateTimeOffset = DateTime.Now.AddMonths(-MonthsInPast);
                }

                if (RestrictMonthsInFuture)
                {
                    request.TimeMaxDateTimeOffset = DateTime.Now.AddMonths(MonthsInFuture);
                }

                Log.Information("Processing single updates.");

                Events response;
                var gone_error = false;
                var modified_error = false;

                do
                {
                    //TODO (obelix30) - convert to Polly after retargeting to 4.5
                    try
                    {
                        response = await request.ExecuteAsync(cancellationToken);
                    }
                    catch (Google.GoogleApiException ex)
                    {
                        if (GoogleServices.IsTransientError(ex.HttpStatusCode, ex.Error))
                        {
                            await Task.Delay(TimeSpan.FromMinutes(10), cancellationToken);
                            response = await request.ExecuteAsync(cancellationToken);
                        }
                        else
                        {
                            throw;
                        }
                    }

                    foreach (var a in response.Items)
                    {
                        if (a.Id != null)
                        {
                            try
                            {
                                if (deleteGoogleAppointments)
                                {
                                    if (a.Status != "cancelled")
                                    {
                                        await GoogleEventsResource.Delete(SyncAppointmentsGoogleFolder, a.Id).ExecuteAsync(cancellationToken);
                                    }
                                }
                                else if (a.ExtendedProperties != null && (a.ExtendedProperties.Private__ != null && a.ExtendedProperties.Private__.ContainsKey(key) || a.ExtendedProperties.Shared != null && a.ExtendedProperties.Shared.ContainsKey(key)))
                                {
                                    AppointmentPropertiesUtils.ResetGoogleOutlookId(a);
                                    if (a.Status != "cancelled")
                                    {
                                        await GoogleEventsResource.Update(a, SyncAppointmentsGoogleFolder, a.Id).ExecuteAsync(cancellationToken);
                                    }
                                }
                            }
                            catch (Google.GoogleApiException ex)
                            {
                                if (ex.HttpStatusCode == HttpStatusCode.Gone)
                                {
                                    gone_error = true;
                                }
                                else if (ex.HttpStatusCode == HttpStatusCode.PreconditionFailed)
                                {
                                    modified_error = true;
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }
                    request.PageToken = response.NextPageToken;
                }
                while (!string.IsNullOrEmpty(request.PageToken));

                if (modified_error)
                {
                    Log.Debug("Some Google appointments modified before update.");
                }
                if (gone_error)
                {
                    Log.Debug("Some Google appointments gone before deletion.");
                }
                return gone_error || modified_error;
            }
            catch (WebException ex)
            {
                throw new Google.GoogleApiException(service, message, ex);  //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }
            catch (NullReferenceException ex)
            {
                throw new Google.GoogleApiException(service, message, new WebException("Error accessing feed", ex));    //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }
        }

        /// <summary>
        /// Resets Google appointment matches via batch updates.
        /// </summary>
        /// <param name="deleteGoogleAppointments">Should Google appointments be updated or deleted.</param>        
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>If error occured.</returns>
        internal async Task<bool> BatchResetGoogleAppointmentMatches(bool deleteGoogleAppointments, CancellationToken cancellationToken)
        {
            const string message = "Error updating Google appointments.";
            const string service = "GCSM.Synchronizer.BatchResetGoogleAppointmentMatches";

            var key = OutlookPropertiesUtils.GetKey();

            try
            {
                var request = GoogleEventsResource.List(SyncAppointmentsGoogleFolder);
                request.PageToken = null;

                if (RestrictMonthsInPast)
                {
                    request.TimeMinDateTimeOffset = DateTime.Now.AddMonths(-MonthsInPast);
                }

                if (RestrictMonthsInFuture)
                {
                    request.TimeMaxDateTimeOffset = DateTime.Now.AddMonths(MonthsInFuture);
                }

                Log.Information("Processing batch updates.");

                Events response;
                var br = new BatchRequest(GoogleCalendarService);

                var events = new Dictionary<string, Event>();
                var gone_error = false;
                var modified_error = false;
                var rate_error = false;
                var current_batch_rate_error = false;

                var batches = 1;
                do
                {

                    //TODO (obelix30) - check why sometimes exception happen like below,  we have custom backoff attached
                    //                    Google.GoogleApiException occurred
                    //User Rate Limit Exceeded[403]
                    //Errors[
                    //    Message[User Rate Limit Exceeded] Location[- ] Reason[userRateLimitExceeded] Domain[usageLimits]

                    //TODO (obelix30) - convert to Polly after retargeting to 4.5
                    try
                    {
                        response = await request.ExecuteAsync(cancellationToken);
                    }
                    catch (Google.GoogleApiException ex)
                    {
                        if (GoogleServices.IsTransientError(ex.HttpStatusCode, ex.Error))
                        {
                            await Task.Delay(TimeSpan.FromMinutes(10), cancellationToken);
                            response = await request.ExecuteAsync(cancellationToken);
                        }
                        else
                        {
                            throw;
                        }
                    }

                    foreach (var a in response.Items)
                    {
                        if (a.Id != null && !events.ContainsKey(a.Id))
                        {
                            IClientServiceRequest r = null;
                            if (a.Status != "cancelled")
                            {
                                if (deleteGoogleAppointments)
                                {
                                    events.Add(a.Id, a);
                                    r = GoogleEventsResource.Delete(SyncAppointmentsGoogleFolder, a.Id);

                                }
                                else if (a.ExtendedProperties != null && (a.ExtendedProperties.Private__ != null && a.ExtendedProperties.Private__.ContainsKey(key) || a.ExtendedProperties.Shared != null && a.ExtendedProperties.Shared.ContainsKey(key)))
                                {
                                    events.Add(a.Id, a);
                                    AppointmentPropertiesUtils.ResetGoogleOutlookId(a);
                                    r = GoogleEventsResource.Update(a, SyncAppointmentsGoogleFolder, a.Id);
                                }
                            }

                            if (r != null)
                            {
                                br.Queue<Event>(r, (content, error, ii, msg) =>
                                {
                                    if (error != null && msg != null)
                                    {
                                        if (msg.StatusCode == HttpStatusCode.PreconditionFailed)
                                        {
                                            modified_error = true;
                                        }
                                        else if (msg.StatusCode == HttpStatusCode.Gone)
                                        {
                                            gone_error = true;
                                        }
                                        else if (GoogleServices.IsTransientError(msg.StatusCode, error))
                                        {
                                            rate_error = true;
                                            current_batch_rate_error = true;
                                        }
                                        else
                                        {
                                            Log.Information($"Batch error: {error}");
                                        }
                                    }
                                });
                                if (br.Count >= GoogleServices.BatchRequestSize)
                                {
                                    if (current_batch_rate_error)
                                    {
                                        current_batch_rate_error = false;
                                        await Task.Delay(GoogleServices.BatchRequestBackoffDelay);
                                        Log.Debug($"Back-Off waited {GoogleServices.BatchRequestBackoffDelay}ms before next retry...");

                                    }
                                    await br.ExecuteAsync(cancellationToken);
                                    // TODO(obelix30): https://github.com/google/google-api-dotnet-client/issues/725
                                    br = new BatchRequest(GoogleCalendarService);

                                    Log.Information($"Batch of Google changes finished ({batches})");
                                    batches++;
                                }
                            }
                        }
                    }
                    request.PageToken = response.NextPageToken;
                }
                while (!string.IsNullOrEmpty(request.PageToken));

                if (br.Count > 0)
                {
                    await br.ExecuteAsync(cancellationToken);
                    Log.Information($"Batch of Google changes finished ({batches})");
                }
                if (modified_error)
                {
                    Log.Debug("Some Google appointment modified before update.");
                }
                if (gone_error)
                {
                    Log.Debug("Some Google appointment gone before deletion.");
                }
                if (rate_error)
                {
                    Log.Debug("Rate errors received.");
                }

                return gone_error || modified_error || rate_error;
            }
            catch (WebException ex)
            {
                throw new Google.GoogleApiException(service, message, ex);  //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }
            catch (NullReferenceException ex)
            {
                throw new Google.GoogleApiException(service, message, new WebException("Error accessing feed", ex));    //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            }
        }

        public Event GetGoogleAppointment(string gid)
        {
            var ga = GetGoogleAppointmentById(gid);

            if (ga != null)
            {
                return ga;
            }
            else
            {
                var policy = registry.Get<RetryPolicy>("Standard");

                var result = policy.ExecuteAndCapture(() =>
                {
                    return GoogleEventsResource.Get(SyncAppointmentsGoogleFolder, gid).Execute();
                });

                return result.Result;
            }
        }

        public void DeleteGoogleAppointment(Event ga)
        {
            if (ga != null && !ga.Status.Equals("cancelled"))
            {
                GoogleEventsResource.Delete(SyncAppointmentsGoogleFolder, ga.Id).Execute();
            }
        }

        public EventsResource.InstancesRequest GetGoogleAppointmentInstances(string id)
        {
            return GoogleEventsResource.Instances(SyncAppointmentsGoogleFolder, id);
        }


        /// <summary>
        /// Load the contacts from Google and Outlook
        /// </summary>
        public void LoadContacts()
        {
            LoadOutlookContacts();
            LoadGoogleGroups();
            LoadGoogleContacts();
            RemoveOutlookDuplicatedContacts();
            RemoveGoogleDuplicatedContacts();
        }
        private void LoadGoogleContacts()
        {
            LoadGoogleContacts(null);
            Log.Debug($"Google Contacts Found: {GoogleContacts.Count}");
        }

        public bool IsOutlookAppointmentToBeProcessed(AppointmentItem oa)
        {
            try
            {
                if (oa == null)
                {
                    Log.Debug("Outlook Appointment was null ==> skipping");
                    return false;
                }

                /*if (oa.IsDeleted())
                {
                    Log.Debug($"Outlook Appointment {oa.ToLogString()} was deleted ==> skipping");
                    return false;
                }*/

                if (string.IsNullOrEmpty(oa.Subject) && oa.Start == AppointmentSync.outlookDateInvalid)
                {
                    Log.Debug($"Outlook Appointment {oa.ToLogString()} didn't have subject or start date ==> skipping");
                    return false;
                }

                if (oa.MeetingStatus == OlMeetingStatus.olMeetingCanceled || oa.MeetingStatus == OlMeetingStatus.olMeetingReceivedAndCanceled)
                {
                    Log.Debug($"Outlook Appointment {oa.ToLogString()} canceled ==> skipping");
                    return false;
                }

                if (RestrictMonthsInPast)
                {
                    if (oa.IsRecurring)
                    {
                        RecurrencePattern rp = null;

                        try
                        {
                            rp = oa.GetRecurrence();
                            if (rp.PatternEndDate < DateTime.Now.AddMonths(-MonthsInPast))
                            {
                                Log.Debug($"Outlook Appointment {oa.ToLogString()} recurrence pattern ended before the sync range (MonthsInPast = {MonthsInPast}) ==> skipping");
                                return false;
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Debug(ex, $"Exception getting Outlook Appointment {oa.ToLogString()} recurrence pattern ==> Skipping: {ex.Message}");
                            oa.ToDebugLog();
                            return false;
                        }
                        finally
                        {
                            if (rp != null)
                            {
                                Marshal.ReleaseComObject(rp);
                            }
                        }
                    }
                    else
                    {
                        if (oa.End < DateTime.Now.AddMonths(-MonthsInPast))
                        {
                            Log.Debug($"Outlook Appointment {oa.ToLogString()} ended before the sync range (MonthsInPast = {MonthsInPast}) ==> skipping");
                            return false;
                        }
                    }
                }

                if (RestrictMonthsInFuture)
                {
                    if (oa.IsRecurring)
                    {
                        RecurrencePattern rp = null;

                        try
                        {
                            rp = oa.GetRecurrence();
                            if (rp.PatternStartDate > DateTime.Now.AddMonths(MonthsInFuture))
                            {
                                Log.Debug($"Outlook Appointment {oa.ToLogString()} starts after the sync range (MonthsInFuture = {MonthsInFuture}) ==> skipping");
                                return false;
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Warning($"Exception getting Outlook appointment {oa.ToLogString()} recurrence pattern ==> Skipping: {ex.Message}");
                            oa.ToDebugLog();
                            return false;
                        }
                        finally
                        {
                            if (rp != null)
                            {
                                Marshal.ReleaseComObject(rp);
                            }
                        }
                    }
                    else
                    {
                        if (oa.Start > DateTime.Now.AddMonths(MonthsInFuture))
                        {
                            Log.Debug($"Outlook Appointment starts after the sync range (MonthsInFuture = {MonthsInFuture}) ==> skipping");
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //this is needed because some appointments throw exceptions
                Log.Warning($"Exception getting Outlook Appointment {oa.ToLogString()} ==> Skipping: {ex.Message}");
                Log.Debug(ex, "Exception");
                oa.ToDebugLog();
                return false;
            }

            return true;
        }

        /// <summary>
        /// Remove duplicates from Google: two different Google appointments pointing to the same Outlook appointment.
        /// </summary>
        private void RemoveGoogleDuplicatedAppointments()
        {
            Log.Information("Removing Google duplicated appointments...");

            var appointments = new Dictionary<string, int>();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Log.Debug($"Total allocated memory: {GC.GetTotalMemory(false):N0}");

            //scan all Google appointments
            for (var i = 0; i < GoogleAppointments.Count; i++)
            {
                var ga1 = GoogleAppointments[i];
                if (ga1 == null)
                {
                    continue;
                }

                try
                {
                    var oid = AppointmentPropertiesUtils.GetGoogleOutlookId(ga1);

                    //check if Google event is linked to Outlook appointment
                    if (string.IsNullOrEmpty(oid))
                    {
                        continue;
                    }

                    //check if there is already another Google event linked to the same Outlook appointment 
                    if (appointments.ContainsKey(oid))
                    {
                        var ga2 = GoogleAppointments[appointments[oid]];
                        if (ga2 == null)
                        {
                            appointments.Remove(oid);
                            continue;
                        }
                        else if (ga2 == ga1 || ga2.Id == ga1.Id ||
                            ga2.RecurringEventId != null && ga1.RecurringEventId != null && ga2.RecurringEventId == ga1.RecurringEventId ||
                            ga2.RecurringEventId == ga1.Id || ga1.RecurringEventId == ga2.Id)
                        { //Do nothing, same event found or at least different recurrences of the same recurrence master
                            continue;
                        }

                        var oa = GetOutlookAppointmentById(oid);

                        if (IsOutlookAppointmentToBeProcessed(oa))
                        {
                            var gid = AppointmentPropertiesUtils.GetOutlookGoogleId(oa);

                            //check to which Outlook appoinment Google event is linked
                            if (AppointmentPropertiesUtils.GetGoogleId(ga1) == gid)
                            {
                                AppointmentPropertiesUtils.ResetGoogleOutlookId(ga2);
                                Log.Debug($"Duplicated appointment: {ga2.ToLogString()}.");
                                appointments[oid] = i;
                            }
                            else if (AppointmentPropertiesUtils.GetGoogleId(ga2) == gid)
                            {
                                AppointmentPropertiesUtils.ResetGoogleOutlookId(ga1);
                                Log.Debug($"Duplicated appointment: {ga1.ToLogString()}.");
                            }
                            else
                            {
                                AppointmentPropertiesUtils.ResetGoogleOutlookId(ga1);
                                AppointmentPropertiesUtils.ResetGoogleOutlookId(ga2);
                                AppointmentPropertiesUtils.ResetOutlookGoogleId(oa);
                                Save(ref oa);
                            }
                        }
                        else
                        {
                            //duplicated Google events found, but Outlook appointment does not exist
                            //so lets clean the link from Google events  
                            AppointmentPropertiesUtils.ResetGoogleOutlookId(ga1);
                            AppointmentPropertiesUtils.ResetGoogleOutlookId(ga2);
                            appointments.Remove(oid);
                        }

                        if (oa != null)
                        {
                            Marshal.ReleaseComObject(oa);
                        }
                    }
                    else
                    {
                        appointments.Add(oid, i);
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some appointments throw exceptions
                    Log.Debug($"Accessing Google appointment: {ga1.ToLogString()} threw an exception. Skipping: {ex.Message}");
                    continue;
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Log.Debug($"Total allocated memory: {GC.GetTotalMemory(false):N0}");
        }

        /// <summary>
        /// Remove duplicates from Outlook: two different Outlook appointments pointing to the same Google appointment.
        /// Such situation typically happens when copy/paste'ing synchronized appointment in Outlook
        /// </summary>
        private void RemoveOutlookDuplicatedAppointments()
        {
            Log.Information("Removing Outlook duplicated appointments...");

            var appointments = new Dictionary<string, string>();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Log.Debug($"Total allocated memory: {GC.GetTotalMemory(false):N0}");

            Items items = null;

            try
            {
                items = GetAppointmentItems();

                //scan all appointments
                for (var i = 1; i <= items.Count; i++)
                {
                    AppointmentItem oa1 = null;

                    try
                    {
                        oa1 = items[i] as AppointmentItem;

                        if (!IsOutlookAppointmentToBeProcessed(oa1))
                        {
                            if (oa1 != null)
                            {
                                Marshal.ReleaseComObject(oa1);
                            }
                            continue;
                        }

                        var gid = AppointmentPropertiesUtils.GetOutlookGoogleId(oa1);
                        //check if Outlook appointment is linked to Google event
                        if (string.IsNullOrEmpty(gid))
                        {
                            Marshal.ReleaseComObject(oa1);
                            continue;
                        }

                        //check if there is already another Outlook appointment linked to the same Google event 
                        if (appointments.ContainsKey(gid))
                        {
                            var oid2 = appointments[gid];

                            if (string.IsNullOrEmpty(oid2))
                            {
                                Marshal.ReleaseComObject(oa1);
                                continue;
                            }

                            var o = OutlookNameSpace.GetItemFromID(oid2);
                            //"is" operator creates an implicit variable (COM leak), so unfortunately we need to avoid pattern matching
#pragma warning disable IDE0019 // Use pattern matching
                            var oa2 = o as AppointmentItem;
#pragma warning restore IDE0019 // Use pattern matching
                            if (oa2 == null)
                            {
                                appointments.Remove(gid);
                                Marshal.ReleaseComObject(oa1);
                                continue;
                            }

                            var ga = GetGoogleAppointmentById(gid);
                            if (ga != null)
                            {
                                var oid = AppointmentPropertiesUtils.GetGoogleOutlookId(ga);
                                //check to which Outlook appoinment Google event is linked
                                if (AppointmentPropertiesUtils.GetOutlookId(oa1) == oid)
                                {
                                    AppointmentPropertiesUtils.ResetOutlookGoogleId(oa2);
                                    if (!string.IsNullOrEmpty(oa2.Subject))
                                    {
                                        Log.Debug($"Duplicated appointment: {oa2.ToLogString()}.");
                                    }
                                    appointments[gid] = AppointmentPropertiesUtils.GetOutlookId(oa1);
                                    Save(ref oa2);
                                }
                                else if (AppointmentPropertiesUtils.GetOutlookId(oa2) == oid)
                                {
                                    AppointmentPropertiesUtils.ResetOutlookGoogleId(oa1);
                                    if (!string.IsNullOrEmpty(oa1.Subject))
                                    {
                                        Log.Debug($"Duplicated appointment: {oa1.ToLogString()}.");
                                    }
                                    Save(ref oa1);
                                }
                                else
                                {
                                    //duplicated Outlook appointments found, but Google event does not exist
                                    //so lets clean the link from Outlook appointments  
                                    AppointmentPropertiesUtils.ResetOutlookGoogleId(oa1);
                                    AppointmentPropertiesUtils.ResetOutlookGoogleId(oa2);
                                    appointments.Remove(gid);
                                    Save(ref oa1);
                                    Save(ref oa2);
                                }
                            }
                            Marshal.ReleaseComObject(oa2);
                        }
                        else
                        {
                            appointments.Add(gid, AppointmentPropertiesUtils.GetOutlookId(oa1));
                        }
                        Marshal.ReleaseComObject(oa1);
                    }
                    catch (Exception ex)
                    {
                        //this is needed because some appointments throw exceptions
                        if (oa1 != null && !string.IsNullOrEmpty(oa1.Subject))
                        {
                            Log.Warning($"Accessing Outlook appointment: {oa1.ToLogString()} internal . Skipping: {ex.Message}");
                            oa1.ToDebugLog();
                        }
                        else
                        {
                            Log.Warning($"Accessing Outlook appointment internal . Skipping: {ex.Message}");
                        }
                        Log.Debug(ex, "Exception");

                        continue;
                    }
                }
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                Log.Debug($"Total allocated memory: {GC.GetTotalMemory(false):N0}");
            }
        }

        /// <summary>
        /// Remove duplicates from Google: two different Google contacts pointing to the same Outlook contact.
        /// </summary>
        private void RemoveGoogleDuplicatedContacts()
        {
            Log.Information("Removing Google duplicated contacts...");
            Log.Debug($"GoogleContacts before removing duplicates: {GoogleContacts.Count}");
            var contacts = new Dictionary<string, int>();

            Log.Debug($"DEBUG VERSION 4 +++ DEBUG VERSION 4 +++");
            //scan all Google contacts
            for (var i = GoogleContacts.Count - 1; i >= 0; i--)
            {
                Log.Debug($"[{i}/{GoogleContacts.Count}]...");
                Log.Debug($"Accessing GoogeContacts[i]...");
                var c1 = GoogleContacts[i];
                Log.Debug($"Accessed GoogeContacts[i]...");
                if (c1 == null)
                {
                    continue;
                }

                var googleUniqueIdentifierName = string.Empty;
                try
                {
                    Log.Debug($"ContactPropertiesUtils.GetGoogleUniqueIdentifierName(c1)...");
                    googleUniqueIdentifierName = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(c1);
                    Log.Debug($"Accessed ContactPropertiesUtils.GetGoogleUniqueIdentifierName(c1)...");
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "Exception");
                }

                Log.Debug($"[{i}/{GoogleContacts.Count}] Checking contact {googleUniqueIdentifierName} for duplicates...");

                try
                {
                    var oid = ContactPropertiesUtils.GetGoogleOutlookContactId(c1);
                    //check if Google contact is linked to Outlook contact
                    if (string.IsNullOrEmpty(oid))
                    {
                        Log.Debug($"[{i}/{GoogleContacts.Count}] Checked contact {googleUniqueIdentifierName} for duplicates (not yet linked to Outlook)...");
                        continue;
                    }

                    //check if there is already another Google contact linked to the same Outlook contact 
                    if (contacts.ContainsKey(oid))
                    {
                        var index = contacts[oid];
                        if (index >= 0 && index < GoogleContacts.Count)
                        {
                            Log.Debug($"[{i}/{GoogleContacts.Count}] Checked contact {googleUniqueIdentifierName} for duplicates (duplicate index out of range, removing from sync)");
                            contacts.Remove(oid);
                            continue;
                        }

                        var c2 = GoogleContacts[contacts[oid]];
                        if (c2 == null)
                        {
                            Log.Debug($"[{i}/{GoogleContacts.Count}] Checked contact {googleUniqueIdentifierName} for duplicates (duplicate not found (null), removing from sync)");
                            contacts.Remove(oid);
                            continue;
                        }

                        if (c1.ResourceName == c2.ResourceName)
                        { //same contact, remove from GoogleContacts
                            Log.Debug($"[{i}/{GoogleContacts.Count}] Checked contact {googleUniqueIdentifierName} for duplicates, duplicate found but same resourceName, removing from sync...");
                            GoogleContacts.RemoveAt(contacts[oid]);
                            continue;
                        }

                        var a = GetOutlookContactById(oid);
                        if (a != null)
                        {
                            var gid = ContactPropertiesUtils.GetOutlookGoogleId(a);
                            //check to which Outlook contact Google contact is linked
                            if (ContactPropertiesUtils.GetGoogleId(c1) == gid)
                            {
                                var googleUniqueIdentifierName2 = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(c2);
                                ContactPropertiesUtils.ResetGoogleOutlookId(c2);
                                Log.Debug($"[{i}/{GoogleContacts.Count}] Duplicated contact found: {googleUniqueIdentifierName2}, reset link.");
                                contacts[oid] = i;
                            }
                            else if (ContactPropertiesUtils.GetGoogleId(c2) == gid)
                            {

                                ContactPropertiesUtils.ResetGoogleOutlookId(c1);
                                Log.Debug($"[{i}/{GoogleContacts.Count}] Duplicated contact found: {googleUniqueIdentifierName}, reset link.");
                            }
                            else
                            {
                                ContactPropertiesUtils.ResetGoogleOutlookId(c1);
                                ContactPropertiesUtils.ResetGoogleOutlookId(c2);
                                ContactPropertiesUtils.ResetOutlookGoogleId(this, a);
                                Log.Debug($"[{i}/{GoogleContacts.Count}] Duplicated contact found: {googleUniqueIdentifierName}, reset link on both sides.");
                            }
                        }
                        else
                        {
                            //duplicated Google contacts found, but Outlook contact does not exist
                            //so lets clean the link from Google contacts
                            ContactPropertiesUtils.ResetGoogleOutlookId(c1);
                            ContactPropertiesUtils.ResetGoogleOutlookId(c2);
                            contacts.Remove(oid);
                            Log.Debug($"[{i}/{GoogleContacts.Count}] Duplicated Google contact found, but Outlook contact does not exist: {googleUniqueIdentifierName}, clean the link from Google contacts");
                        }
                    }
                    else
                    {
                        Log.Debug($"[{i}/{GoogleContacts.Count}] Checked contact {googleUniqueIdentifierName} for duplicates , no duplicate found...");
                        contacts.Add(oid, i);
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some contacts throw exceptions
                    if (c1 != null)
                    {
                        Log.Warning($"Exception Accessing Google contact: {googleUniqueIdentifierName} internal. Skipping: {ex.Message}");
                    }
                    else
                    {
                        Log.Warning($"Exception Accessing Google contact internal (was null). Skipping: {ex.Message}");
                    }
                    Log.Debug(ex, "Exception");
                    c1.ToDebugLog();
                    continue;
                }
            }

            Log.Debug("Removed Google duplicated contacts.");
            Log.Debug($"GoogleContacts after removing duplicates: {GoogleContacts.Count}");
        }

        /// <summary>
        /// Remove duplicates from Outlook: two different Outlook contacts pointing to the same Google contact.
        /// Such situation typically happens when copy/paste'ing synchronized contact in Outlook
        /// </summary>
        private void RemoveOutlookDuplicatedContacts()
        {
            Log.Information("Removing Outlook duplicated contacts...");

            var contacts = new Dictionary<string, int>();

            //scan all contacts
            for (var i = 1; i <= OutlookContacts.Count; i++)
            {
                ContactItem olc1 = null;

                try
                {
                    olc1 = OutlookContacts[i] as ContactItem;
                    if (olc1 == null)
                    {
                        continue;
                    }

                    var gid = ContactPropertiesUtils.GetOutlookGoogleId(olc1);
                    //check if Outlook contact  is linked to Google contact
                    if (string.IsNullOrEmpty(gid))
                    {
                        continue;
                    }

                    //check if there is already another Outlook contact linked to the same Google contact 
                    if (contacts.ContainsKey(gid))
                    {
                        var o = OutlookContacts[contacts[gid]];
                        //"is" operator creates an implicit variable (COM leak), so unfortunately we need to avoid pattern matching
#pragma warning disable IDE0019 // Use pattern matching
                        var olc2 = o as ContactItem;
#pragma warning restore IDE0019 // Use pattern matching
                        if (olc2 == null)
                        {
                            contacts.Remove(gid);
                            continue;
                        }

                        var c = GetGoogleContactById(gid);
                        if (c != null)
                        {
                            var oid = ContactPropertiesUtils.GetGoogleOutlookContactId(c);
                            //check to which Outlook contact Google contact is linked
                            if (!string.IsNullOrEmpty(oid) && oid.Equals(ContactPropertiesUtils.GetOutlookId(olc1), StringComparison.InvariantCultureIgnoreCase))
                            {
                                ContactPropertiesUtils.ResetOutlookGoogleId(this, olc2);
                                Log.Debug($"Duplicated contact: {olc2.ToLogString()}.");
                                contacts[oid] = i;
                            }
                            else if (!string.IsNullOrEmpty(oid) && oid.Equals(ContactPropertiesUtils.GetOutlookId(olc2), StringComparison.InvariantCultureIgnoreCase))
                            {
                                ContactPropertiesUtils.ResetOutlookGoogleId(this, olc1);
                                Log.Debug($"Duplicated contact: {olc1.ToLogString()}.");
                            }
                            else
                            {
                                //duplicated Outlook contacts found, but Google contact does not exist
                                //so lets clean the link from Outlook contacts  
                                ContactPropertiesUtils.ResetOutlookGoogleId(this, olc1);
                                ContactPropertiesUtils.ResetOutlookGoogleId(this, olc2);
                                contacts.Remove(gid);
                            }
                        }
                        else
                        {
                            //duplicated Outlook contacts found, but Google contact does not exist
                            //so lets clean the link from Outlook contacts
                            ContactPropertiesUtils.ResetOutlookGoogleId(this, olc1);
                            ContactPropertiesUtils.ResetOutlookGoogleId(this, olc2);
                            contacts.Remove(gid);
                        }
                    }
                    else
                    {
                        contacts.Add(gid, i);
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some contacts throw exceptions
                    Log.Debug($"Accessing Outlook contact: {olc1.ToLogString()} internal . Skipping: {ex.Message}");
                    continue;
                }
            }
        }

        public void LoadAppointments()
        {
            LoadGoogleAppointments();
            RemoveOutlookDuplicatedAppointments();
            RemoveGoogleDuplicatedAppointments();
            LoadOutlookAppointments();
        }

        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
        public void MatchContacts()
        {
            LoadContacts();

            Log.Debug("Matching Contacts...");
            Contacts = ContactsMatcher.MatchContacts(this, out var duplicateDataException);
            if (duplicateDataException != null)
            {
                if (DuplicatesFound != null)
                {
                    DuplicatesFound("Google duplicates found", duplicateDataException.Message);
                }
                else
                {
                    Log.Warning(duplicateDataException.Message);
                }
            }
            Log.Debug($"Person Matches Found: {Contacts.Count}");
        }

        /// <summary>
        /// Load the appointments from Google and Outlook and match them
        /// </summary>
        public void MatchAppointments()
        {
            LoadAppointments();
            Appointments = AppointmentsMatcher.MatchAppointments(this);
            Log.Debug($"Appointment Matches Found: {Appointments.Count}");
        }

        private void LogSyncParams()
        {
            Log.Debug("Synchronization options:");
            Log.Debug($"Profile: {SyncProfile}");
            Log.Debug($"SyncOption: {SyncOption}");
            Log.Debug($"SyncDelete: {SyncDelete}");
            Log.Debug($"PromptDelete: {PromptDelete}");

            if (SyncContacts)
            {
                Log.Debug("Sync contacts");
                if (_OutlookNameSpace != null)
                {
                    var fld = _OutlookNameSpace.GetFolderFromID(SyncContactsFolder);
                    Log.Debug($"SyncContactsFolder: {fld.FullFolderPath}");
                }

                Log.Debug($"SyncContactsForceRTF: {SyncContactsForceRTF}");
                Log.Debug($"SyncPhotos: {SyncPhotos}");
                Log.Debug($"UseFileAs: {UseFileAs}");
            }

            if (SyncAppointments)
            {
                try
                {
                    Log.Debug("Sync appointments");
                    Log.Debug($"MonthsInPast: {MonthsInPast}");
                    Log.Debug($"MonthsInFuture: {MonthsInFuture}");
                    if (_OutlookNameSpace != null)
                    {
                        var fld = _OutlookNameSpace.GetFolderFromID(SyncAppointmentsFolder);
                        Log.Debug($"SyncAppointmentsFolder: {fld.FullFolderPath}");
                    }
                    Log.Debug($"SyncAppointmentsGoogleFolder: {SyncAppointmentsGoogleFolder}");
                    Log.Debug($"SyncAppointmentsForceRTF: {SyncAppointmentsForceRTF}");
                    Log.Debug($"SyncAppointmentsPrviate: {SyncAppointmentsPrivate}");
                }
                catch (COMException ex)
                {
                    Log.Debug(ex, "Exception");
                    LogoffOutlook();
                    LoginToOutlook();
                }
            }
        }

        public void Sync()
        {
            lock (_syncRoot)
            {
                try
                {
                    if (string.IsNullOrEmpty(SyncProfile))
                    {
                        Log.Error("Must set a sync profile. This should be different on each user/computer you sync on.");
                        return;
                    }

                    LogSyncParams();

                    SyncedCount = 0;
                    DeletedCount = 0;
                    ErrorCount = 0;
                    SkippedCount = 0;
                    SkippedCountNotMatches = 0;
                    ConflictResolution = ConflictResolution.Cancel;
                    DeleteGoogleResolution = DeleteResolution.Cancel;
                    DeleteOutlookResolution = DeleteResolution.Cancel;

                    if (SyncContacts)
                    {
                        MatchContacts();
                    }

                    if (SyncAppointments)
                    {
                        Log.Information($"Outlook default time zone: {TimeZoneInfo.Local.Id}");
                        Log.Information($"Google default time zone: {SyncAppointmentsGoogleTimeZone}");
                        if (string.IsNullOrEmpty(Timezone))
                        {
                            TimeZoneChanges?.Invoke(SyncAppointmentsGoogleTimeZone);
                            Log.Information("Timezone not configured, changing to default value from Google, it could be adjusted later in GUI.");
                        }
                        else if (string.IsNullOrEmpty(SyncAppointmentsGoogleTimeZone))
                        {
                            //Timezone was set, but some users do not have time zone set in Google
                            SyncAppointmentsGoogleTimeZone = Timezone;
                        }
                        MappingBetweenTimeZonesRequired = false;
                        if (TimeZoneInfo.Local.Id != AppointmentSync.IanaToWindows(SyncAppointmentsGoogleTimeZone))
                        {
                            MappingBetweenTimeZonesRequired = true;
                            Log.Warning($"Different time zones in Outlook ({TimeZoneInfo.Local.Id}) and Google (mapped to {AppointmentSync.IanaToWindows(SyncAppointmentsGoogleTimeZone)})");
                        }
                        MatchAppointments();
                    }

                    if (SyncContacts)
                    {
                        if (Contacts == null)
                        {
                            return;
                        }

                        TotalCount = Contacts.Count + SkippedCountNotMatches;

                        //Resolve Google duplicates from matches to be synced
                        ResolveDuplicateContacts(GoogleContactDuplicates);

                        //Remove Outlook duplicates from matches to be synced
                        if (OutlookContactDuplicates != null)
                        {
                            for (var i = OutlookContactDuplicates.Count - 1; i >= 0; i--)
                            {
                                var match = OutlookContactDuplicates[i];
                                if (Contacts.Contains(match))
                                {
                                    //ToDo: If there has been a resolution for a duplicate above, there is still skipped increased, check how to distinguish
                                    SkippedCount++;
                                    Contacts.Remove(match);
                                }
                            }
                        }

                        Log.Information("Syncing groups...");
                        ContactsMatcher.SyncGroups(this);

                        Log.Information("Syncing contacts...");
                        ContactsMatcher.SyncContacts(this);

                        SaveContacts(Contacts);
                    }

                    if (SyncAppointments)
                    {
                        if (Appointments == null)
                        {
                            return;
                        }

                        TotalCount += Appointments.Count + SkippedCountNotMatches;

                        Log.Information("Syncing appointments...");
                        AppointmentsMatcher.SyncAppointments(this);

                        DeleteAppointments(Appointments);
                    }
                }
                finally
                {
                    GoogleContacts = null;
                    GoogleAppointments = null;
                    OutlookContactDuplicates = null;
                    GoogleContactDuplicates = null;
                    GoogleGroups = null;
                    Contacts = null;
                    Appointments = null;
                }
            }
        }

        private void ResolveDuplicateContacts(Collection<ContactMatch> googleContactDuplicates)
        {
            if (googleContactDuplicates != null)
            {
                for (var i = googleContactDuplicates.Count - 1; i >= 0; i--)
                {
                    ResolveDuplicateContact(googleContactDuplicates[i]);
                }
            }
        }

        private void ResolveDuplicateContact(ContactMatch match)
        {
            if (Contacts.Contains(match))
            {
                if (SyncOption == SyncOption.MergePrompt)
                {
                    //For each OutlookDuplicate: Ask user for the GoogleContact to be synced with
                    for (var j = match.AllOutlookContactMatches.Count - 1; j >= 0 && match.AllGoogleContactMatches.Count > 0; j--)
                    {
                        var olci = match.AllOutlookContactMatches[j];
                        var oc = olci.GetOriginalItemFromOutlook();

                        using (var r = new ConflictResolver())
                        {
                            switch (r.ResolveDuplicate(olci, match.AllGoogleContactMatches, out var googleContact))
                            {
                                case ConflictResolution.Skip:
                                case ConflictResolution.SkipAlways: //Keep both entries and sync it to both sides
                                    match.AllGoogleContactMatches.Remove(googleContact);
                                    match.AllOutlookContactMatches.Remove(olci);
                                    Contacts.Add(new ContactMatch(null, googleContact));
                                    Contacts.Add(new ContactMatch(olci, null));
                                    break;
                                case ConflictResolution.OutlookWins:
                                case ConflictResolution.OutlookWinsAlways: //Keep Outlook and overwrite Google
                                    match.AllGoogleContactMatches.Remove(googleContact);
                                    match.AllOutlookContactMatches.Remove(olci);
                                    UpdateContact(oc, googleContact, match);
                                    SaveContact(new ContactMatch(olci, googleContact));
                                    break;
                                case ConflictResolution.GoogleWins:
                                case ConflictResolution.GoogleWinsAlways: //Keep Google and overwrite Outlook
                                    match.AllGoogleContactMatches.Remove(googleContact);
                                    match.AllOutlookContactMatches.Remove(olci);
                                    UpdateContact(googleContact, oc, match.GoogleContactDirty, match.matchedById);
                                    SaveContact(new ContactMatch(olci, googleContact));
                                    break;
                                default:
                                    throw new ApplicationException("Cancelled");
                            }
                        }

                        //Cleanup the match, i.e. assign a proper OutlookContact and GoogleContact, because can be deleted before
                        match.OutlookContact = match.AllOutlookContactMatches.Count == 0 ? null : match.AllOutlookContactMatches[0];
                    }
                }

                //Cleanup the match, i.e. assign a proper OutlookContact and GoogleContact, because can be deleted before
                match.GoogleContact = match.AllGoogleContactMatches.Count == 0 ? null : match.AllGoogleContactMatches[0];

                if (match.AllOutlookContactMatches.Count == 0)
                {
                    //If all OutlookContacts have been assigned by the users ==> Create one match for each remaining Google Person to sync them to Outlook
                    Contacts.Remove(match);
                    foreach (var googleContact in match.AllGoogleContactMatches)
                    {
                        Contacts.Add(new ContactMatch(null, googleContact));
                    }
                }
                else if (match.AllGoogleContactMatches.Count == 0)
                {
                    //If all GoogleContacts have been assigned by the users ==> Create one match for each remaining Outlook Contact to sync them to Google
                    Contacts.Remove(match);
                    foreach (var outlookContact in match.AllOutlookContactMatches)
                    {
                        Contacts.Add(new ContactMatch(outlookContact, null));
                    }
                }
                else
                {
                    SkippedCount++;
                    Contacts.Remove(match);
                }
            }
        }

        public void DeleteAppointments(List<AppointmentMatch> appointments)
        {
            foreach (var match in appointments)
            {
                try
                {
                    DeleteAppointment(match);
                }
                catch (Exception ex)
                {
                    if (ErrorEncountered != null)
                    {
                        ErrorCount++;
                        SyncedCount--;
                        var s = match.OutlookAppointment != null ? match.OutlookAppointment.ToLogString() + ")" : match.GoogleAppointment.ToLogString();
                        var message = $"Failed to synchronize appointment: {s}:\n{ex.Message}";
                        var newEx = new Exception(message, ex);
                        ErrorEncountered("Error", newEx);
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        private void DeleteAppointmentNoGoogle(AppointmentMatch match)
        {
            var oa = match.OutlookAppointment;
            var gid = AppointmentPropertiesUtils.GetOutlookGoogleId(oa);
            var name = oa.ToLogString();
            if (!string.IsNullOrEmpty(gid))
            {// There was a sync before, but the Google Appointment was deleted
                if (SyncOption == SyncOption.OutlookToGoogleOnly)
                {
                    SkippedCount++;
                    Log.Debug($"Skipped deletion of Outlook appointment because of SyncOption {SyncOption}: {name}.");
                    try
                    {
                        AppointmentPropertiesUtils.ResetOutlookGoogleId(oa);
                        Save(ref oa);
                    }
                    catch (Exception)
                    {
                        Log.Warning($"Error resetting match for Outlook appointment: {name}.");
                    }
                }
                else if (!SyncDelete)
                {
                    SkippedCount++;
                    Log.Debug($"Skipped deletion of Outlook appointment because SyncDeletion is switched off: {name}.");
                }
                else
                {

                    if (AppointmentsMatcher.RecipientsCount(oa) > 1)  //slave to be updated has some recipients ==> Don't delete, recreate from outlook to google
                                                                                     //&& AppointmentPropertiesUtils.GetSyncId(googleAppointment, GetPartnerAppointmentsFolder(googleAppointment)) != null  //slave to be updated was snychronized before, i.e. not freshly created                        
                    {
                        //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment   
                        switch (SyncOption)
                        {
                            case SyncOption.MergeOutlookWins:
                            case SyncOption.OutlookToGoogleOnly:
                                //overwrite Google appointment
                                Log.Debug($"Multiple attendees found on Outlook appointment, invitation maybe NOT sent by Google. Outlook appointment is recreating Google because of SyncOption {SyncOption}: {oa.ToLogString()}.");
                                RecreateGoogleAppointment(ref match);
                                break;
                            case SyncOption.MergeGoogleWins:
                                //overwrite Google appointment
                                Log.Debug($"Multiple attendees found on Outlook appointment, invitation maybe NOT sent by Google. Outlook appointment is recreating Google, even though SyncOption {SyncOption}: {oa.ToLogString()}.");
                                RecreateGoogleAppointment(ref match);
                                break;
                            case SyncOption.GoogleToOutlookOnly:
                                //skip overwrite Google appointment
                                Log.Debug($"Multiple attendees found on Outlook appointment, invitation maybe NOT sent by Google, skipping Google to recreate from Outlook, even though SyncOption {SyncOption}: {oa.ToLogString()}.");
                                SkippedCount++; //updated = true;
                                break;
                            case SyncOption.MergePrompt:
                                //prompt for sync option
                                if (ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                    ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                    ConflictResolution != ConflictResolution.SkipAlways)
                                {
                                    using (var r = new ConflictResolver())
                                    {
                                        ConflictResolution = r.Resolve($"Cannot delete Outlook appointment because multiple participants found, invitation maybe NOT sent by Google:\"{oa.ToLogString()}\". Do you want to recreate it back from Outlook to Google?", oa, null, this);
                                    }
                                }
                                switch (ConflictResolution)
                                {
                                    case ConflictResolution.Skip:
                                    case ConflictResolution.SkipAlways: //Skip
                                    case ConflictResolution.GoogleWins:
                                    case ConflictResolution.GoogleWinsAlways: //Keep Google without update back
                                        SkippedCount++;
                                        Log.Debug($"{ConflictResolution}: skipped recreating appointment from Outlook to Google because multiple participants found: \"{oa.ToLogString()}\".");
                                        break;
                                    case ConflictResolution.OutlookWins:
                                    case ConflictResolution.OutlookWinsAlways: //Keep Outlook and overwrite Google    
                                        Log.Debug($"{ConflictResolution}: recreated appointment from Outlook to Google because multiple participants found: \"{oa.ToLogString()}\".");
                                        RecreateGoogleAppointment(ref match);
                                        break;
                                    default:
                                        throw new ApplicationException("Cancelled");
                                }
                                break;
                        }
                    }
                    else
                    {

                        // Google appointment was deleted, delete outlook appointment
                        try
                        {
                            //First reset OutlookGoogleContactId to restore it later from trash
                            AppointmentPropertiesUtils.ResetOutlookGoogleId(oa);
                            Save(ref oa);
                        }
                        catch (Exception)
                        {
                            Log.Warning($"Error resetting match for Outlook appointment: {name}.");
                        }

                        oa.Delete();

                        DeletedCount++;
                        Log.Information($"Deleted Outlook appointment: {name}.");
                    }
                }
            }
            else
            {//There was no sync before
                if (SyncOption == SyncOption.GoogleToOutlookOnly)
                {
                    SkippedCount++;
                    Log.Debug("Skipped Update from Outlook to NEW Google appointment because of SyncOption " + SyncOption + ":" + name + ".");
                }
            }
        }

        private void DeleteAppointmentNoOutlook(AppointmentMatch match)
        {
            var ga = match.GoogleAppointment;
            var name = ga.ToLogString();
            var oid = AppointmentPropertiesUtils.GetGoogleOutlookId(ga);
            if (!string.IsNullOrEmpty(oid))
            {// There was a sync before, but the Outlook Appointment was deleted
                if (SyncOption == SyncOption.GoogleToOutlookOnly)
                {
                    SkippedCount++;
                    Log.Debug($"Skipped deletion of Google appointment because of SyncOption {SyncOption}: {name}.");
                    if (ga.Status != "cancelled")
                    {
                        try
                        {
                            AppointmentPropertiesUtils.ResetGoogleOutlookId(ga);
                            SaveGoogleAppointment(ga);
                        }
                        catch (Exception)
                        {
                            Log.Warning($"Error resetting match for Google appointment: {name}.");
                        }
                    }
                }
                else if (!SyncDelete)
                {
                    SkippedCount++;
                    Log.Debug($"Skipped deletion of Google appointment because SyncDeletion is switched off: {name}.");
                }
                else if (ga.Status != "cancelled")
                {

                    if (ga.Creator != null && !AppointmentSync.IsOrganizer(ga.Creator.Email) || ga.Attendees != null && ga.Attendees.Count > 1)                     
                    {
                        switch (SyncOption)
                        {
                            case SyncOption.MergeGoogleWins:
                            case SyncOption.GoogleToOutlookOnly:
                                //overwrite Outlook appointment                            
                                Log.Debug($"Multiple attendees found on Google, invitation maybe NOT sent by Outlook. Google appointment is recreating Outlook because of SyncOption {SyncOption}: {ga.ToLogString()}.");
                                RecreateOutlookAppointment(ref match);
                                break;
                            case SyncOption.MergeOutlookWins:
                                //overwrite Outlook appointment
                                Log.Debug($"Multiple attendees found on Google, invitation maybe NOT sent by  Outlook. Google appointment is recreating Outlook, even though SyncOption {SyncOption}: {ga.ToLogString()}.");
                                RecreateOutlookAppointment(ref match);
                                break;
                            case SyncOption.OutlookToGoogleOnly:
                                //skip overwrite Google appointment
                                Log.Debug($"Multiple attendees found on Google, invitation maybe NOT sent by Outlook, skipping Outlook appointment recreating Google, even though SyncOption {SyncOption}: {ga.ToLogString()}.");
                                SkippedCount++; //updated = true;
                                break;
                            case SyncOption.MergePrompt:
                                //promp for sync option
                                if (
                                    //ConflictResolution != ConflictResolution.OutlookWinsAlways && //Shouldn't be used, because Google seems to be the master of the appointment
                                    ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                    ConflictResolution != ConflictResolution.SkipAlways)
                                {
                                    using (var r = new ConflictResolver())
                                    {
                                        ConflictResolution = r.Resolve($"Cannot delete Google appointment because multiple attendees found, invitation maybe NOT sent by Outlook: \"{ga.ToLogString()}\". Do you want to recreate it back from Google to Outlook?", null, ga, this);
                                    }
                                }
                                switch (ConflictResolution)
                                {
                                    case ConflictResolution.Skip:
                                    case ConflictResolution.SkipAlways: //Skip
                                    case ConflictResolution.OutlookWins:
                                    case ConflictResolution.OutlookWinsAlways: //Keep Outlook    
                                        SkippedCount++;
                                        Log.Debug($"{ConflictResolution}:Skipped recreating appointment from Outlook to Google because multiple attendees found on Google: \"{ga.ToLogString()}\"");
                                        break;
                                    case ConflictResolution.GoogleWins:
                                    case ConflictResolution.GoogleWinsAlways: //Keep Google and overwrite Outlook                           
                                        Log.Debug($"{ConflictResolution}: \"{ga.ToLogString()}\".");
                                        RecreateOutlookAppointment(ref match);
                                        break;
                                    default:
                                        throw new ApplicationException("Cancelled");
                                }

                                break;
                        }
                    }
                    else
                    {

                        GoogleEventsResource.Delete(SyncAppointmentsGoogleFolder, ga.Id).Execute();
                        DeletedCount++;
                        Log.Information($"Deleted Google appointment: {name}.");
                    }
                }

            }
            else
            {//There was no sync before
                if (SyncOption == SyncOption.OutlookToGoogleOnly)
                {
                    SkippedCount++;
                    Log.Debug("Skipped Update from Gootle to NEW Outlook appointment because of SyncOption " + SyncOption + ":" + name + ".");
                }
            }
        }

        public void RecreateOutlookAppointment(ref AppointmentMatch match)
        {
            match.OutlookAppointment = CreateOutlookAppointmentItem(SyncAppointmentsFolder);
            UpdateAppointment(ref match.GoogleAppointment, ref match.OutlookAppointment, match.GoogleAppointmentExceptions);            
        }

        public void RecreateGoogleAppointment(ref AppointmentMatch match)
        {
            match.GoogleAppointment = Factory.NewEvent();
            UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment, ref match.GoogleAppointmentExceptions);
        }

        public void DeleteAppointment(AppointmentMatch match)
        {
            if (match.GoogleAppointment == null && match.OutlookAppointment != null)
            {
                DeleteAppointmentNoGoogle(match);
            }
            else if (match.GoogleAppointment != null && match.OutlookAppointment == null)
            {
                DeleteAppointmentNoOutlook(match);
            }
        }

        public void SaveContacts(List<ContactMatch> contacts)
        {
            foreach (var match in contacts)
            {
                try
                {
                    SaveContact(match);
                }
                catch (Exception ex)
                {
                    if (ErrorEncountered != null)
                    {
                        ErrorCount++;
                        SyncedCount--;
                        var s = match.OutlookContact != null ? match.OutlookContact.FileAs : ContactPropertiesUtils.GetGoogleUniqueIdentifierName(match.GoogleContact);
                        var message = $"Failed to synchronize contact: {s}. \nPlease check the contact, if any Email already exists on Google contacts side or if there is too much or invalid data in the notes field. \nIf the problem persists, please try recreating the contact or report the error:\n{ex.Message}";
                        var newEx = new Exception(message, ex);
                        ErrorEncountered("Error", newEx);
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        public bool SaveContact(ContactMatch match)
        {
            if (match.GoogleContact != null && match.OutlookContact != null)
            {
                if (match.GoogleContactDirty)
                {
                    //google contact was modified. save.
                    if (SaveGoogleContact(match))
                    {
                        SyncedCount++;
                        Log.Information($"Updated contact from Outlook to Google: \"{match}\".");
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            else if (match.GoogleContact == null && match.OutlookContact != null)
            {   //Delete OutlookContact, but only if SyncDelete and not SyncOption.OutlookToGoogleOnly
                var name = match.OutlookContact.FileAs;
                if (SyncOption == SyncOption.OutlookToGoogleOnly)
                {
                    SkippedCount++;
                    Log.Debug($"Skipped Deletion of Outlook contact because of SyncOption {SyncOption}: {name}.");
                }
                else if (!SyncDelete)
                {
                    SkippedCount++;
                    Log.Debug($"Skipped Deletion of Outlook contact because SyncDeletion is switched off: {name}.");
                }
                else if ((match.OutlookContact.UserProperties.GoogleContactId != null || this.SyncOption == SyncOption.GoogleToOutlookOnly)
                     && this.DeleteOutlookResolution != DeleteResolution.KeepOutlook
                     && this.DeleteOutlookResolution != DeleteResolution.KeepOutlookAlways)
                {

                    // peer google contact was deleted, delete outlook contact
                    var item = match.OutlookContact.GetOriginalItemFromOutlook();
                    try
                    {
                        //First reset OutlookGoogleContactId to restore it later from trash
                        ContactPropertiesUtils.ResetOutlookGoogleId(this, item);
                        Save(ref item);
                    }
                    catch (Exception)
                    {
                        Log.Warning($"Error resetting match for Outlook contact: \"{name}\".");
                    }

                    item.Delete();
                    DeletedCount++;
                    Log.Information($"Deleted Outlook contact: \"{name}\".");

                }
                else
                {
                    SkippedCount++;
                    Log.Debug($"Skipped Deletion of Outlook contact (SyncOption {SyncOption}, DeleteResolution.{this.DeleteGoogleResolution}): {name}.");
                }
            }
            else if (match.GoogleContact != null && match.OutlookContact == null)
            {  //Delete GoogleContact, but only if SyncDelete and not SyncOption.GoogleToOutlookOnly
                var name = ContactMatch.GetName(match.GoogleContact);
                if (SyncOption == SyncOption.GoogleToOutlookOnly)
                {
                    SkippedCount++;
                    Log.Debug($"Skipped Deletion of Google contact because of SyncOption {SyncOption}: {name}.");
                }
                else if (!SyncDelete)
                {
                    SkippedCount++;
                    Log.Debug($"Skipped Deletion of Google contact because SyncDeletion is switched off: {name}.");
                }
                else if ((!string.IsNullOrEmpty(ContactPropertiesUtils.GetGoogleOutlookContactId(match.GoogleContact)) || this.SyncOption == SyncOption.OutlookToGoogleOnly)
                     && this.DeleteGoogleResolution != DeleteResolution.KeepGoogle
                     && this.DeleteGoogleResolution != DeleteResolution.KeepGoogleAlways)
                {
                    GooglePeopleResource.DeleteContact(match.GoogleContact.ResourceName).Execute();
                    DeletedCount++;
                    Log.Information($"Deleted Google contact: \"{name}\".");
                }
                else
                {
                    SkippedCount++;
                    Log.Debug($"Skipped Deletion of Google contact (SyncOption {SyncOption}, DeleteResolution.{this.DeleteGoogleResolution}): {name}.");
                }
            }
            else
            {
                throw new ArgumentNullException("To save contacts, at least a GoogleContact or OutlookContact must be present.");
            }

            return true;
        }
        /// <summary>
        /// Only for testing purposes, sets the recurrence exception to null
        /// </summary>
        /// <param name="master"></param>
        /// <param name="slave"></param>
        internal void UpdateAppointment(AppointmentItem master, ref Event slave)
        {
            List<Event> l = null;
            UpdateAppointment(master, ref slave, ref l);
        }

        /// <summary>
        /// Updates Outlook appointment from master to slave (including groups/categories)
        /// </summary>
        public void UpdateAppointment(AppointmentItem master, ref Event slave, ref List<Event> GoogleAppointmentExceptions)
        {
            var updated = false;

            if ((slave.Creator != null && !AppointmentSync.IsOrganizer(slave.Creator.Email) || slave.Attendees != null && slave.Attendees.Count > 1) // && AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(this.SyncProfile, slave) != null)
                && !(AppointmentsMatcher.RecipientsCount(master) > 1 && AppointmentPropertiesUtils.GetOutlookGoogleId(master) != null)) //To avoid endless loop calling recursively UpdateAppointment
            {
                //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment   
                switch (SyncOption)
                {
                    case SyncOption.MergeGoogleWins:
                    case SyncOption.GoogleToOutlookOnly:
                        //overwrite Outlook appointment                            
                        Log.Debug($"Different Organizer or multiple attendees found on Google, invitation maybe NOT sent by Outlook. Google appointment is overwriting Outlook because of SyncOption {SyncOption}: {master.ToLogString()}.");
                        UpdateAppointment(ref slave, ref master, GoogleAppointmentExceptions);
                        break;
                    case SyncOption.MergeOutlookWins:
                        //overwrite Outlook appointment
                        Log.Debug($"Different Organizer or multiple attendees found on Google, invitation maybe NOT sent by Outlook. Google appointment is overwriting Outlook, even though SyncOption {SyncOption}: {master.ToLogString()}.");
                        UpdateAppointment(ref slave, ref master, GoogleAppointmentExceptions);
                        break;
                    case SyncOption.OutlookToGoogleOnly:
                        //skip overwrite Google appointment
                        Log.Debug($"Different Organizer or multiple attendees found on Google, invitation maybe NOT sent by Outlook, skipping Outlook appointment overwrite Google, even though SyncOption {SyncOption}: {master.ToLogString()}.");
                        SkippedCount++; //updated = true;
                        break;
                    case SyncOption.MergePrompt:
                        //promp for sync option
                        if (
                            //ConflictResolution != ConflictResolution.OutlookWinsAlways && //Shouldn't be used, because Google seems to be the master of the appointment
                            ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                            ConflictResolution != ConflictResolution.SkipAlways)
                        {
                            using (var r = new ConflictResolver())
                            {
                                ConflictResolution = r.Resolve($"Cannot update appointment from Outlook to Google because different Organizer or multiple attendees found on Google, invitation maybe NOT sent by Outlook: \"{master.ToLogString()}\". Do you want to update it back from Google to Outlook?", slave, master, this);
                            }
                        }
                        switch (ConflictResolution)
                        {
                            case ConflictResolution.Skip:
                            case ConflictResolution.SkipAlways: //Skip
                            case ConflictResolution.OutlookWins:
                            case ConflictResolution.OutlookWinsAlways: //Keep Outlook    
                                SkippedCount++;
                                Log.Information($"{ConflictResolution}:Skipped updating appointment from Outlook to Google because different organizer found on Google: \"{master.ToLogString()}\". Google organizer is " + slave.Creator.Email.Trim().ToLower().Replace("@googlemail.", "@gmail.") + " and user name is " + UserName.Trim().ToLower().Replace("@googlemail.", "@gmail.") + ".");
                                break;
                            case ConflictResolution.GoogleWins:
                            case ConflictResolution.GoogleWinsAlways: //Keep Google and overwrite Outlook                           
                                Log.Debug($"{ConflictResolution}: \"{master.ToLogString()}\".");
                                UpdateAppointment(ref slave, ref master, null);
                                break;
                            default:
                                throw new ApplicationException("Cancelled");
                        }

                        break;
                }
            }
            else //Only update, if invitation was not sent on Google side or freshly created during this sync  
            {
                updated = true;
            }

            if (updated)
            {
                //match.GoogleContactDirty = true;
                AppointmentSync.UpdateAppointment(master, slave);

                if (slave.Creator == null || AppointmentSync.IsOrganizer(slave.Creator.Email))
                {
                    AppointmentPropertiesUtils.SetGoogleOutlookId(slave, master);
                    var updatedSlave = SaveGoogleAppointment(slave);

                    updated = updatedSlave != null && (updatedSlave != slave || updatedSlave.ETag != slave.ETag);
                    if (updated)
                        slave = updatedSlave;
                }
                else
                    updated = false;

                if (updated)
                {
                    AppointmentPropertiesUtils.SetOutlookGoogleId(master, slave);
                    Save(ref master);
                }

                //After saving Google Appointment => also sync recurrence exceptions and save again
                //TODO (obelix30),  create test for birthdays (auto created by gmail, so user is not organizer)
                //and check what happens if recurrence exception is provoked
                if (updated && slave.Creator != null && !AppointmentSync.IsOrganizer(slave.Creator.Email))
                    Log.Debug($"Could not update an appointment from Outlook to Google (you are not organizer, invitation maybe NOT sent by Outlook): \"{slave.ToLogString()}\".");
                else if (updated && master.IsRecurring && master.RecurrenceState == OlRecurrenceState.olApptMaster)
                {
                    Log.Debug($"Updated appointment master from Outlook to Google, now updating the recurrence exceptions: \"{slave.ToLogString()}\".");

                    if (AppointmentSync.UpdateRecurrenceExceptions(master, slave, ref GoogleAppointmentExceptions, this))
                    {
                        var updatedSlave = SaveGoogleAppointment(slave);

                        updated = updatedSlave != null && (updatedSlave != slave || updatedSlave.ETag != slave.ETag);
                        if (updated)
                            slave = updatedSlave;
                    }
                    else
                        updated = false;

                    if (updated) //still updated
                        Log.Debug($"Updated appointment recurrence instances from Outlook to Google: \"{slave.ToLogString()}\".");
                    else
                        Log.Debug($"Could not update all appointment recurrence instances from Outlook to Google: \"{slave.ToLogString()}\".");
                }

                if (updated)
                {
                    SyncedCount++;
                    Log.Information($"Updated appointment from Outlook to Google: \"{master.ToLogString()}\".");
                }


            }
        }

        private static bool Save(ref AppointmentItem oa)
        {
            return Save(ref oa, false);
        }

        private static bool Save(ref AppointmentItem oa, bool forced)
        {
            if (!oa.Saved || forced)
            {
                try
                { //Try to save 2 times, because sometimes the first save fails with a COMException (Outlook aborted)
                    oa.Save();
                }
                catch (ArgumentException ex)
                {
                    Log.Warning($"Error saving Outlook Appointment {oa.ToLogString()}\nException: {ex.Message}");
                    Log.Debug(ex, "Exception");
                    if (ex.ParamName != null)
                    {
                        Log.Debug($"Invalid param: {ex.ParamName}");
                    }
                    oa.ToDebugLog();
                    throw new ApplicationException($"Error saving Outlook Appointment { oa.ToLogString() }\nException: { ex.Message}", ex);
                }
                catch (Exception)
                {
                    try
                    {
                        oa.Save();
                    }
                    catch (COMException ex)
                    {
                        Log.Warning($"Error saving Outlook appointment {oa.ToLogString()}\nException: {ex.Message}");
                        Log.Debug(ex, "Exception");
                        oa.ToDebugLog();
                        return false;
                    }
                }
            }
            return true;
        }


        private static bool Save(ref ContactItem oc)
        {
            if (!oc.Saved)
            {
                try
                { //Try to save 3 times, because sometimes the first save fails with a COMException (Object not found)
                    oc.Save();
                }
                catch (ArgumentException ex)
                {
                    Log.Warning($"Error saving Outlook Contact {oc.ToLogString()}\nException: {ex.Message}");
                    Log.Debug(ex, "Exception");
                    if (ex.ParamName != null)
                    {
                        Log.Debug($"Invalid param: {ex.ParamName}");
                    }
                    //oc.ToDebugLog();
                    throw new ApplicationException($"Error saving Outlook Contact { oc.ToLogString() }\nException: { ex.Message}", ex);
                }
                catch (Exception)
                {
                    try
                    {
                        oc.Save();
                    }

                    catch (COMException ex)
                    {
                        Log.Warning($"Error saving Outlook contact {oc.ToLogString()}\nException: {ex.Message}");
                        Log.Debug(ex, "Exception");
                        //oc.ToDebugLog();
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Updates Outlook appointment from master to slave (including groups/categories)
        /// </summary>
        public bool UpdateAppointment(ref Event master, ref AppointmentItem slave, List<Event> googleAppointmentExceptions)
        {
            var updated = false;

            if (AppointmentsMatcher.RecipientsCount(slave) > 1 && AppointmentPropertiesUtils.GetOutlookGoogleId(slave) != null
              && !(master.Creator != null && !AppointmentSync.IsOrganizer(master.Creator.Email) || master.Attendees != null && master.Attendees.Count > 1))  //To avoid endless loop calling recursively UpdateAppointment
            {
                switch (SyncOption)
                {
                    case SyncOption.MergeOutlookWins:
                    case SyncOption.OutlookToGoogleOnly:
                        //overwrite Google appointment
                        Log.Debug($"Multiple attendees found on Outlook, invitation maybe NOT sent by Google. Outlook appointment is overwriting Google because of SyncOption {SyncOption}: {master.ToLogString()}.");
                        UpdateAppointment(slave, ref master, ref googleAppointmentExceptions);
                        break;
                    case SyncOption.MergeGoogleWins:
                        //overwrite Google appointment
                        Log.Debug($"Multiple attendees found on Outlook, invitation maybe NOT sent by Google. Outlook appointment is overwriting Google, even though SyncOption {SyncOption}: {master.ToLogString()}.");
                        UpdateAppointment(slave, ref master, ref googleAppointmentExceptions);
                        break;
                    case SyncOption.GoogleToOutlookOnly:
                        //skip overwrite Google appointment
                        Log.Debug($"Multiple attendees found on Outlook, invitation maybe NOT sent by Google, skipping Google appointment to overwrite Outlook, even though SyncOption {SyncOption}: {master.ToLogString()}.");
                        SkippedCount++; //updated = true;
                        break;
                    case SyncOption.MergePrompt:
                        //promp for sync option
                        if (ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                            ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                            ConflictResolution != ConflictResolution.SkipAlways)
                        {
                            using (var r = new ConflictResolver())
                            {
                                ConflictResolution = r.Resolve($"Cannot update appointment from Google to Outlook because multiple participants found: \"{master.ToLogString()}\". Do you want to update it back from Outlook to Google?", slave, master, this);
                            }
                        }
                        switch (ConflictResolution)
                        {
                            case ConflictResolution.Skip:
                            case ConflictResolution.SkipAlways: //Skip
                            case ConflictResolution.GoogleWins:
                            case ConflictResolution.GoogleWinsAlways: //Keep Google without update back
                                SkippedCount++;
                                Log.Debug($"{ConflictResolution}: skipped updating appointment from Google to Outlook because multiple participants found: \"{master.ToLogString()}\".");
                                break;
                            case ConflictResolution.OutlookWins:
                            case ConflictResolution.OutlookWinsAlways: //Keep Outlook and overwrite Google    
                                Log.Debug($"{ConflictResolution}: updated appointment from Outlook to Google because multiple participants found: \"{master.ToLogString()}\".");
                                UpdateAppointment(slave, ref master, ref googleAppointmentExceptions);
                                break;
                            default:
                                throw new ApplicationException("Cancelled");
                        }
                        break;
                }
            }
            else //Only update, if invitation was not sent on Outlook side or freshly created during this sync  
            {
                updated = true;
            }

            if (updated)
            {
                if (AppointmentSync.UpdateAppointment(master, slave))
                {
                    AppointmentPropertiesUtils.SetOutlookGoogleId(slave, master);

                    if (!Save(ref slave, true))
                    {
                        return false;
                    }

                    AppointmentPropertiesUtils.SetGoogleOutlookId(master, slave);
                    master = SaveGoogleAppointment(master);

                    //SyncedCount++;
                    //Log.Debug($"Updated appointment master from Google to Outlook, now updating the recurrence exceptions: \"{master.ToLogString()}\".");

                    //After saving Outlook Appointment => also sync recurrence exceptions and increase SyncCount
                    if (master.Recurrence != null && googleAppointmentExceptions != null)
                    {
                        Log.Debug($"Updated appointment master from Google to Outlook, now updating the recurrence exceptions: \"{master.ToLogString()}\".");

                        updated = AppointmentSync.UpdateRecurrenceExceptions(googleAppointmentExceptions, ref slave, this);                        

                        if (updated) //still updated
                            Log.Debug($"Updated appointment recurrence instances from Google to Outlook: \"{master.ToLogString()}\".");
                        else
                            Log.Debug($"Could not update all appointment recurrence instances from Google to Outlook: \"{master.ToLogString()}\".");
                    }
                }
                else
                {
                    updated = false;
                    SkippedCount++;

                    var gid = AppointmentPropertiesUtils.GetOutlookGoogleId(slave);
                    if (!string.IsNullOrWhiteSpace(gid))
                    {
                        AppointmentPropertiesUtils.ResetOutlookGoogleId(slave);
                        if (!Save(ref slave))
                        {
                            return false;
                        }
                    }

                    var oid = AppointmentPropertiesUtils.GetGoogleOutlookId(master);
                    if (!string.IsNullOrWhiteSpace(oid))
                    {
                        AppointmentPropertiesUtils.ResetGoogleOutlookId(master);
                        master = SaveGoogleAppointment(master);
                    }
                }

                if (updated)
                {
                    SyncedCount++;
                    Log.Information($"Updated appointment from Google to Outlook: \"{master.ToLogString()}\".");
                }
            }



            return true;
        }

        private void SaveOutlookContact(ref Person gc, ContactItem oc)
        {
            ContactPropertiesUtils.SetOutlookGoogleId(oc, gc);
            Save(ref oc);
            ContactPropertiesUtils.SetGoogleOutlookId(gc, oc);
            var gc1 = SaveGoogleContact(gc);
            if (gc1 != null)
            {
                gc = gc1;
                ContactPropertiesUtils.SetOutlookGoogleId(oc, gc);
                Save(ref oc);
                if (Synchronizer.SyncPhotos)
                {
                    SaveOutlookPhoto(gc, oc);
                }
            }
            else
            {
                ContactPropertiesUtils.ResetOutlookGoogleId(this, oc);
                Save(ref oc);
            }
        }

        public bool SaveGoogleContact(ContactMatch match)
        {
            var oc = match.OutlookContact.GetOriginalItemFromOutlook();

            ContactPropertiesUtils.SetGoogleOutlookId(match.GoogleContact, oc);
            match.GoogleContact = SaveGoogleContact(match.GoogleContact);

            if (match.GoogleContact != null)
            {
                ContactPropertiesUtils.SetOutlookGoogleId(oc, match.GoogleContact);
                Save(ref oc);
                //Now save the Photo
                if (Synchronizer.SyncPhotos)
                {
                    SaveGooglePhoto(match, oc);
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        //private static string GetXml(Person gc)
        //{
        //    using (var ms = new MemoryStream())
        //    {
        //        gc.ContactEntry.SaveToXml(ms);
        //        var sr = new StreamReader(ms);
        //        ms.Seek(0, SeekOrigin.Begin);
        //        return sr.ReadToEnd();
        //    }
        //}

        private Person InsertGoogleContact(Person gc)
        {
            //insert contact.
            //var feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

            try
            {
                var policyWrap = registryWrapPolicies.Get<PolicyWrap>("Contact Write");

                var result = policyWrap.ExecuteAndCapture(() =>
                {
                    return GooglePeopleResource.CreateContact(gc).Execute();
                });

                return result.Result;
            }
            catch (Google.GoogleApiException ex) when (ex.Error != null && ex.Error.ErrorResponseContent.Contains("Resource has been exhausted (e.g. check quota)")) //ToDo: Check counterpart of GDataRequestException in Google People Api (is it really GoogleApiException?)
            {
                var bio = ContactPropertiesUtils.GetGoogleBiographyValue(gc);
                Log.Warning($"Skipping contact {gc.ToLogString()}, it has too large notes field: {bio.Length} characters. Please shorten notes in Outlook contact, otherwise you risk loosing information stored there.");
                return null; //ToDo: Check, what happens if returned null? Maybe one reason for the intermittently deleted Outlook contacts?
            }
            catch (Google.GoogleApiException ex)
            {
                var responseString = ex.Error != null ? System.Web.HttpUtility.HtmlDecode(ex.Error.ErrorResponseContent) : "NoResponseContent";
                Log.Debug(ex, $"ResponseString: {responseString}");
                gc.ToDebugLog();
                var newEx = $"Error saving NEW Google contact:\n{ex.Message}";
                throw new ApplicationException(newEx, ex);
            }
            catch (ApplicationException)
            {//Application already handled internally, no additonal log
                throw;
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                gc.ToDebugLog();
                var newEx = $"Error saving NEW Google contact:\n{ex.Message}";
                throw new ApplicationException(newEx, ex);
            }
        }

        private Person UpdateGoogleContact(Person gc)
        {
            //contact already present in google. just update
            UpdateEmptyUserProperties(gc);
            UpdateExtendedProperties(gc);

            try
            {
                var updateRequest = GooglePeopleResource.UpdateContact(gc, gc.ResourceName);
                updateRequest.UpdatePersonFields = Synchronizer.UPDATE_PERSON_FIELDS;

                var policyWrap = registryWrapPolicies.Get<PolicyWrap>("Contact Write");

                var result = policyWrap.ExecuteAndCapture(() =>
                {
                    return updateRequest.Execute();
                });

                return result.Result;
            }
            catch (ApplicationException)
            {//Application already handled internally, no additonal log
                throw;
            }
            catch (Google.GoogleApiException ex) when (ex.Error != null && ex.Error.ErrorResponseContent.Contains("Resource has been exhausted (e.g. check quota)")) //ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
            {//ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
                var bio = ContactPropertiesUtils.GetGoogleBiographyValue(gc);
                Log.Warning($"Skipping contact {gc.ToLogString()}, it has too large notes field: {bio.Length} characters. Please shorten notes in Outlook contact, otherwise you risk loosing information stored there.");
                return null;
            }
            catch (Google.GoogleApiException ex) when (ex.Error != null && ex.Error.ErrorResponseContent.Contains("Invalid country code: ZZ"))
            {//ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
                Log.Warning($"Skipping contact {gc.ToLogString()}, it has invalid value in country code. Please recreate contact at Google, otherwise you risk loosing information stored there.");
                return null;
            }
            catch (Google.GoogleApiException ex) when (ex.Error != null && ex.Error.ErrorResponseContent.Contains("extendedProperty count limit exceeded: 10"))
            {//ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?
                //some contacts despite having less extendedProperties still can throw such exception
                Log.Debug($"{gc.ToLogString()}: too many extended properties exception thrown: {gc.ClientData.Count}");
                UpdateTooManyExtendedProperties(gc, true);
                return UpdateGoogleContact(gc); //ToDo: Check, maybe endless loop? Maybe one reason for the performance issues reported?
            }
            catch (Google.GoogleApiException ex)
            {//ToDo: Check counterpart of GDataRequestException in Google People Api, really GoogleApiException?                
                var responseString = ex.Error != null ? System.Web.HttpUtility.HtmlDecode(ex.Error.ErrorResponseContent) : "NoErrorResponseContent";
                Log.Debug(ex, $"ResponseString: {responseString}");
                gc.ToDebugLog();
                var newEx = $"Error saving EXISTING Google contact:\n{ex.Message}";
                throw new ApplicationException(newEx, ex);
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                gc.ToDebugLog();
                var newEx = $"Error saving EXISTING Google contact:\n{ex.Message}";
                throw new ApplicationException(newEx, ex);
            }
        }

        /// <summary>
        /// Only save the google contact without photo update
        /// </summary>
        /// <param name="gc"></param>
        internal Person SaveGoogleContact(Person gc)
        {
            //check if this contact was not yet inserted on google.
            if (string.IsNullOrEmpty(gc.ResourceName))  //ToDo: Check (maybe also >0?, earlier it was ContactPropertiesUtils.GetGoogleId(gc)
            {
                return InsertGoogleContact(gc);
            }
            else
            {
                return UpdateGoogleContact(gc);
            }
        }

        private void UpdateExtendedProperties(Person gc)
        {
            RemoveTooManyExtendedProperties(gc);
            RemoveTooBigExtendedProperties(gc);
            RemoveDuplicatedExtendedProperties(gc);
            UpdateEmptyExtendedProperties(gc);
            UpdateTooManyExtendedProperties(gc);
            UpdateTooBigExtendedProperties(gc);
            UpdateDuplicatedExtendedProperties(gc);
        }

        private void UpdateDuplicatedExtendedProperties(Person gc)
        {
            DeleteDuplicatedPropertiesForm form = null;
            var googleUniqueIdentifierName = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc);

            try
            {
                var dups = new HashSet<string>();

                foreach (var p in gc.ClientData)
                {
                    if (dups.Contains(p.Key))
                    {
                        Log.Debug($"{googleUniqueIdentifierName}: for extended property {p.Key} duplicates were found.");
                        if (form == null)
                        {
                            form = new DeleteDuplicatedPropertiesForm();
                        }
                        form.AddExtendedProperty(false, p.Key, "");
                    }
                    else
                    {
                        dups.Add(p.Key);
                    }
                }
                if (form == null)
                {
                    return;
                }

                if (ContactExtendedPropertiesToRemoveIfDuplicated != null)
                {
                    foreach (var p in ContactExtendedPropertiesToRemoveIfDuplicated)
                    {
                        form.AddExtendedProperty(true, p, "");
                    }
                }

                form.SortExtendedProperties();

                if (SettingsForm.Instance.ShowDeleteDuplicatedPropertiesForm(form) == DialogResult.OK)
                {
                    var allCheck = form.removeFromAll;

                    if (allCheck)
                    {
                        if (ContactExtendedPropertiesToRemoveIfDuplicated == null)
                        {
                            ContactExtendedPropertiesToRemoveIfDuplicated = new HashSet<string>();
                        }
                        else
                        {
                            ContactExtendedPropertiesToRemoveIfDuplicated.Clear();
                        }
                        Log.Debug($"{googleUniqueIdentifierName}: will clean some extended properties for all contacts.");
                    }
                    else if (ContactExtendedPropertiesToRemoveIfDuplicated != null)
                    {
                        ContactExtendedPropertiesToRemoveIfDuplicated = null;
                        Log.Debug($"{googleUniqueIdentifierName}: will clean some extended properties for this contact.");
                    }

                    foreach (DataGridViewRow r in form.extendedPropertiesRows)
                    {
                        if (Convert.ToBoolean(r.Cells["Selected"].Value))
                        {
                            var key = r.Cells["Key"].Value.ToString();

                            if (allCheck)
                            {
                                ContactExtendedPropertiesToRemoveIfDuplicated.Add(key);
                            }

                            for (var j = gc.ClientData.Count - 1; j >= 0; j--)
                            {
                                if (gc.ClientData[j].Key == key)
                                {
                                    gc.ClientData.RemoveAt(j);
                                }
                            }

                            Log.Debug($"Extended property to remove: {key}");
                        }
                    }
                }
            }
            finally
            {
                if (form != null)
                {
                    form.Dispose();
                }
            }
        }

        private void UpdateTooBigExtendedProperties(Person gc)
        {
            DeleteTooBigPropertiesForm form = null;

            try
            {
                var googleUniqueIdentifierName = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc);

                foreach (var p in gc.ClientData)
                {
                    if (p.Value != null && p.Value.Length > 1012)
                    {
                        Log.Debug($"{googleUniqueIdentifierName}: for extended property {p.Key} size limit exceeded ({p.Value.Length}). Value is: {p.Value}");
                        if (form == null)
                        {
                            form = new DeleteTooBigPropertiesForm();
                        }
                        form.AddExtendedProperty(false, p.Key, p.Value);
                    }
                }
                if (form == null)
                {
                    return;
                }

                if (ContactExtendedPropertiesToRemoveIfTooBig != null)
                {
                    foreach (var p in ContactExtendedPropertiesToRemoveIfTooBig)
                    {
                        form.AddExtendedProperty(true, p, "");
                    }
                }

                form.SortExtendedProperties();

                if (SettingsForm.Instance.ShowDeleteTooBigPropertiesForm(form) == DialogResult.OK)
                {
                    var allCheck = form.removeFromAll;

                    if (allCheck)
                    {
                        if (ContactExtendedPropertiesToRemoveIfTooBig == null)
                        {
                            ContactExtendedPropertiesToRemoveIfTooBig = new HashSet<string>();
                        }
                        else
                        {
                            ContactExtendedPropertiesToRemoveIfTooBig.Clear();
                        }
                        Log.Debug($"{googleUniqueIdentifierName}: will clean some extended properties for all contacts.");
                    }
                    else if (ContactExtendedPropertiesToRemoveIfTooBig != null)
                    {
                        ContactExtendedPropertiesToRemoveIfTooBig = null;
                        Log.Debug($"{googleUniqueIdentifierName}: will clean some extended properties for this contact.");
                    }

                    foreach (DataGridViewRow r in form.extendedPropertiesRows)
                    {
                        if (Convert.ToBoolean(r.Cells["Selected"].Value))
                        {
                            var key = r.Cells["Key"].Value.ToString();

                            if (allCheck)
                            {
                                ContactExtendedPropertiesToRemoveIfTooBig.Add(key);
                            }

                            for (var j = gc.ClientData.Count - 1; j >= 0; j--)
                            {
                                if (gc.ClientData[j].Key == key)
                                {
                                    gc.ClientData.RemoveAt(j);
                                }
                            }

                            Log.Debug($"Extended property to remove: {key}");
                        }
                    }
                }
            }
            finally
            {
                if (form != null)
                {
                    form.Dispose();
                }
            }
        }

        private void UpdateTooManyExtendedProperties(Person gc, bool force = false)
        {
            var googleUniqueIdentifierName = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc);
            if (force || gc.ClientData.Count > 9)
            {
                if (!force)
                {


                    Log.Debug($"{googleUniqueIdentifierName}: too many extended properties {gc.ClientData.Count}");
                }

                var contactKey = OutlookPropertiesUtils.GetKey();

                using (var form = new DeleteTooManyPropertiesForm())
                {
                    foreach (var p in gc.ClientData)
                    {
                        if (p.Key != contactKey)
                        {
                            form.AddExtendedProperty(false, p.Key, p.Value);
                        }
                    }

                    if (ContactExtendedPropertiesToRemoveIfTooMany != null)
                    {
                        foreach (var p in ContactExtendedPropertiesToRemoveIfTooMany)
                        {
                            form.AddExtendedProperty(true, p, "");
                        }
                    }

                    form.SortExtendedProperties();

                    if (SettingsForm.Instance.ShowDeleteTooManyPropertiesForm(form) == DialogResult.OK)
                    {
                        var allCheck = form.removeFromAll;

                        if (allCheck)
                        {
                            if (ContactExtendedPropertiesToRemoveIfTooMany == null)
                            {
                                ContactExtendedPropertiesToRemoveIfTooMany = new HashSet<string>();
                            }
                            else
                            {
                                ContactExtendedPropertiesToRemoveIfTooMany.Clear();
                            }
                            Log.Debug($"{googleUniqueIdentifierName}: will clean some extended properties for all contacts.");
                        }
                        else if (ContactExtendedPropertiesToRemoveIfTooMany != null)
                        {
                            ContactExtendedPropertiesToRemoveIfTooMany = null;
                            Log.Debug($"{googleUniqueIdentifierName}: will clean some extended properties for this contact.");
                        }

                        foreach (DataGridViewRow r in form.extendedPropertiesRows)
                        {
                            if (Convert.ToBoolean(r.Cells["Selected"].Value))
                            {
                                var key = r.Cells["Key"].Value.ToString();

                                if (allCheck)
                                {
                                    ContactExtendedPropertiesToRemoveIfTooMany.Add(key);
                                }

                                for (var i = gc.ClientData.Count - 1; i >= 0; i--)
                                {
                                    if (gc.ClientData[i].Key == key)
                                    {
                                        gc.ClientData.RemoveAt(i);
                                    }
                                }

                                Log.Debug($"Extended property to remove: {key}");
                            }
                        }
                    }
                }
            }
        }

        private static void UpdateEmptyUserProperties(Person gc)
        {
            // User can create an empty label custom field on the web, but when I retrieve, and update, it throws this:
            // Data Request Error Response: [Line 12, Column 44, element gContact:userDefinedField] Missing attribute: &#39;key&#39;
            // Even though I didn't touch it.  So, I will search for empty keys, and give them a simple name.  Better than deleting...
            /*if (gc.ContactEntry == null)
            {
                return;
            }*/

            if (gc.ClientData != null)
            {
                var fieldCount = 0;
                foreach (var userDefinedField in gc.ClientData)
                {
                    fieldCount++;
                    if (string.IsNullOrEmpty(userDefinedField.Key))
                    {
                        userDefinedField.Key = $"UserField{fieldCount}";
                        Log.Debug($"Set key to user defined field to avoid errors: {userDefinedField.Key}");
                    }

                    //similar error with empty values
                    if (string.IsNullOrEmpty(userDefinedField.Value))
                    {
                        userDefinedField.Value = string.Empty; //userDefinedField.Key;//ToDo: Bug in PeopleApi, if ClientData.Remove is not saved into Google and if set to null or String.Empty, it saves the Key into the Value
                        Log.Debug($"Set value same as key to user defined field to avoid errors: {userDefinedField.Value}");
                        //gc.ClientData.RemoveAt(i); //Remove doesn't work with PeopleApi, see above
                        //Log.Debug($"Removed empty user defined field to avoid errors: {userDefinedField.Key}");
                    }
                }
            }

            if (gc.UserDefined != null)
            {
                var fieldCount = 0;
                for (var i = gc.UserDefined.Count - 1; i >= 0; i--)
                {
                    var userDefinedField = gc.UserDefined[i];
                    fieldCount++;
                    if (string.IsNullOrEmpty(userDefinedField.Key))
                    {
                        userDefinedField.Key = $"UserField{fieldCount}";
                        Log.Debug($"Set key to user defined field to avoid errors: {userDefinedField.Key}");
                    }

                    //similar error with empty values
                    if (string.IsNullOrEmpty(userDefinedField.Value))
                    {
                        //userDefinedField.Value = userDefinedField.Key;
                        //Log.Debug($"Set value to user defined field to avoid errors: {userDefinedField.Value}");
                        gc.UserDefined.RemoveAt(i);
                        Log.Debug($"Removed empty user defined field to avoid errors: {userDefinedField.Key}");
                    }
                }
            }
        }

        private static void UpdateEmptyExtendedProperties(Person gc)
        {
            var googleUniqueIdentifierName = ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc);

            foreach (var p in gc.ClientData)
            {
                if (string.IsNullOrEmpty(p.Value))
                {

                    Log.Debug($"{googleUniqueIdentifierName}: empty value for {p.Key}");
                    //if (p.ChildNodes != null)
                    //{
                    //    Log.Debug($"{fileAs.Value}: childNodes count {p.ChildNodes.Count}");
                    //}
                    //else
                    //{
                    p.Value = p.Key;
                    Log.Debug($"{googleUniqueIdentifierName}: set value to extended property to avoid errors {p.Key}");
                    //}
                }
            }
        }

        private void RemoveDuplicatedExtendedProperties(Person gc)
        {
            if (ContactExtendedPropertiesToRemoveIfDuplicated != null)
            {
                for (var i = gc.ClientData.Count - 1; i >= 0; i--)
                {
                    var key = gc.ClientData[i].Key;
                    if (ContactExtendedPropertiesToRemoveIfDuplicated.Contains(key))
                    {
                        Log.Debug($"{ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc)}: removed (duplicate) {key}");
                        gc.ClientData.RemoveAt(i);
                    }
                }
            }
        }

        private void RemoveTooBigExtendedProperties(Person gc)
        {
            if (ContactExtendedPropertiesToRemoveIfTooBig != null)
            {
                if (gc.ClientData == null)
                    for (var i = gc.ClientData.Count - 1; i >= 0; i--)
                    {
                        if (gc.ClientData[i].Value.Length > 1012)
                        {
                            var key = gc.ClientData[i].Key;
                            if (ContactExtendedPropertiesToRemoveIfTooBig.Contains(key))
                            {
                                Log.Debug($"{ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc)}: removed (size) {key}");
                                gc.ClientData.RemoveAt(i);
                            }
                        }
                    }
            }
        }

        private void RemoveTooManyExtendedProperties(Person gc)
        {
            if (ContactExtendedPropertiesToRemoveIfTooMany != null)
            {
                for (var i = gc.ClientData.Count - 1; i >= 0; i--)
                {
                    var key = gc.ClientData[i].Key;
                    if (ContactExtendedPropertiesToRemoveIfTooMany.Contains(key))
                    {
                        Log.Debug($"{ContactPropertiesUtils.GetGoogleUniqueIdentifierName(gc)}: removed (count) {key}");
                        gc.ClientData.RemoveAt(i);
                    }
                }
            }
        }

        private Event InsertGoogleAppointment(Event ga)
        {
            try
            {
                return GoogleEventsResource.Insert(ga, SyncAppointmentsGoogleFolder).Execute();
            }
            catch (Exception ex)
            {
                Log.Warning($"Error saving new Google appointment {ga.ToLogString()}:\nException: {ex.Message}");

                Log.Debug(ex, "Exception");
                ga.ToDebugLog();

                throw new ApplicationException($"Error saving new Google appointment: {ga.ToLogString()}. \n{ex.Message}", ex);
            }
        }

        private Event UpdateGoogleAppointment(Event ga)
        {
            try
            {
                return GoogleEventsResource.Update(ga, SyncAppointmentsGoogleFolder, ga.Id).Execute();
            }
            catch (Google.GoogleApiException ex) when (ex.Error.Code == 412) //PreCondition Failed
            {
                var responseString = ex.Error != null ? System.Web.HttpUtility.HtmlDecode(ex.Error.ErrorResponseContent) : "NoResponseContent";
                Log.Debug(ex, $"Cannot update appointment {ga.ToLogString()} (precondition failed):\n{ex.Message}\nResponseString: {responseString}");
                ga.ToDebugLog();

                //return ga;
                throw new ApplicationException($"Error saving existing Google appointment {ga.ToLogString()} (precondition failed).\n{ex.Message}", ex);
            }
            catch (Google.GoogleApiException ex) when ((ex.Error.Code == 403) && ex.Error.Message.Equals("The operation can only be performed by the organizer of the event."))
            {
                var msg = $"Cannot update appointment {ga.ToLogString()} (you are not organizer)";
                msg += " - Creator: " + (ga.Creator != null ? ga.Creator.Email : "null");
                msg += " - Organizer: " + (ga.Organizer != null ? ga.Organizer.Email : "null");
                Log.Warning(msg);

                var responseString = ex.Error != null ? System.Web.HttpUtility.HtmlDecode(ex.Error.ErrorResponseContent) : "NoResponseContent";
                Log.Debug(ex, $"ResponseString: {responseString}");
                ga.ToDebugLog();

                //return ga;
                throw new ApplicationException(msg, ex);
            }
            catch (Google.GoogleApiException ex) when ((ex.Error.Code == 403) && ex.Error.Errors.Count > 1 && ex.Error.Errors[0].Reason.Equals("forbiddenForNonOrganizer"))
            {
                var msg = $"Cannot update appointment {ga.ToLogString()} (you are not organizer)";
                msg += " - Creator: " + (ga.Creator != null ? ga.Creator.Email : "null");
                msg += " - Organizer: " + (ga.Organizer != null ? ga.Organizer.Email : "null");
                Log.Warning(msg);

                var responseString = ex.Error != null ? System.Web.HttpUtility.HtmlDecode(ex.Error.ErrorResponseContent) : "NoResponseContent";
                Log.Debug(ex, $"ResponseString: {responseString}");
                ga.ToDebugLog();


                //return ga;
                throw new ApplicationException(msg, ex);
            }
            catch (Google.GoogleApiException ex) when ((ex.Error.Code == 403) && ex.Error.Message.Equals("You need to have writer access to this calendar."))
            {
                Log.Warning($"Cannot update appointment {ga.ToLogString()} (no write access)");

                var responseString = ex.Error != null ? System.Web.HttpUtility.HtmlDecode(ex.Error.ErrorResponseContent) : "NoResponseContent";
                Log.Debug(ex, $"ResponseString: {responseString}");
                ga.ToDebugLog();

                //return ga;
                throw new ApplicationException($"Error saving existing Google appointment {ga.ToLogString()} (no write access).\n{ex.Message}", ex);
            }
            catch (Google.GoogleApiException ex)
            {
                Log.Warning($"Error saving existing Google appointment {ga.ToLogString()}.\n{ex.Message}");

                var responseString = ex.Error != null ? System.Web.HttpUtility.HtmlDecode(ex.Error.ErrorResponseContent) : "NoResponseContent";
                Log.Debug(ex, $"ResponseString: {responseString}");
                ga.ToDebugLog();

                //return ga;
                throw new ApplicationException($"Error saving existing Google appointment {ga.ToLogString()}.\n{ex.Message}", ex);
            }
            catch (Exception ex)
            {
                Log.Warning($"Error saving existing Google appointment {ga.ToLogString()}.\n{ex.Message}");

                Log.Debug(ex, "Exception");
                ga.ToDebugLog();

                //return ga;
                throw new ApplicationException($"Error saving existing Google appointment {ga.ToLogString()}.\n{ex.Message}", ex);
            }
        }

        /// <summary>
        /// Save the google Appointment
        /// </summary>
        /// <param name="ga"></param>
        public Event SaveGoogleAppointment(Event ga)
        {
            try
            {
                //check if this contact was not yet inserted on google.
                if (ga.Id == null)
                {
                    return InsertGoogleAppointment(ga);
                }
                else
                {
                    return UpdateGoogleAppointment(ga);
                }
            }
            catch (ApplicationException)
            {
                //ignore, because already handled in sub function, just return back the current ga
                return ga;
            }
        }

        public void SaveGooglePhoto(ContactMatch match, ContactItem oc)
        {
            var hasOutlookPhoto = oc.HasPhoto();

            if (hasOutlookPhoto)
            {
                // add outlook photo to google
                using (var outlookPhoto = oc.GetOutlookPhoto())
                {
                    if (SaveGooglePhoto(match.GoogleContact, outlookPhoto))
                    {
                        //Just save also the Outlook Contact to have the same lastUpdate date as Google
                        ContactPropertiesUtils.SetOutlookGoogleId(oc, match.GoogleContact);
                        Save(ref oc);
                    }
                }
            }
            else
            {
                var hasGooglePhoto = Utilities.HasContactPhoto(match.GoogleContact);
                if (hasGooglePhoto)
                {
                    //Delete Photo on Google side, if no Outlook photo exists
                    try
                    {
                        GooglePeopleResource.DeleteContactPhoto(match.GoogleContact.ResourceName).Execute();
                    }
                    catch (Google.GoogleApiException ex) when (
                            ex.HttpStatusCode == HttpStatusCode.Forbidden ||
                            ex.HttpStatusCode == HttpStatusCode.NotFound)
                    {
                        Log.Error(ex, $"Exception while deleting Google contact photo for id {match.GoogleContact.ResourceName}");

                    }

                    //Just save the Outlook Contact to have the same lastUpdate date as Google
                    ContactPropertiesUtils.SetOutlookGoogleId(oc, match.GoogleContact);
                    Save(ref oc);
                }
            }
        }

        public bool SaveGooglePhoto(Person person, Bitmap photoBitmap)
        {
            if (photoBitmap != null)
            {
                //Try up to several times to overcome Google issue
                const int num_tries = 5;
                for (var retry = 0; retry < num_tries; retry++)
                {
                    try
                    {
                        using (var bmp = new Bitmap(photoBitmap))
                        {
                            //using (var stream = new MemoryStream(Utilities.BitmapToBytes(bmp)))
                            //{
                            var photoReq = new UpdateContactPhotoRequest()
                            {
                                PhotoBytes = Convert.ToBase64String(Utilities.BitmapToBytes(bmp))
                            };

                            GooglePeopleResource.UpdateContactPhoto(photoReq, person.ResourceName).Execute();

                            ////Just save the Outlook Contact to have the same lastUpdate date as Google
                            //ContactPropertiesUtils.SetOutlookGoogleContactId(this, oc, match.GoogleContact);
                            //Save(ref oa);
                            //}
                        }
                        return true; //Exit because photo save succeeded
                    }
                    catch (Google.GoogleApiException ex) when (
                            ex.HttpStatusCode == HttpStatusCode.Forbidden ||
                            ex.HttpStatusCode == HttpStatusCode.NotFound)
                    {
                        Log.Debug(ex, "Exception");
                        //If Google found a picture for a new Google account, it sets it automatically and throws an error, if updating it with the Outlook photo. 
                        //Therefore save it again and try again to save the photo                               
                        if (retry == num_tries - 1)
                        {
                            ErrorHandler.Handle(new Exception($"Photo of contact {ContactPropertiesUtils.GetGoogleUniqueIdentifierName(person)} couldn't be saved after {num_tries} tries, maybe Google found its own photo and doesn't allow updating it", ex));
                        }
                        else
                        {
                            Thread.Sleep(60 * 1000); //sleep 1 minute
                        }
                    }
                }
            }

            return false;
        }

        public void SaveOutlookPhoto(Person gc, ContactItem oc)
        {
            var hasGooglePhoto = Utilities.HasContactPhoto(gc);
            if (hasGooglePhoto)
            {
                // add google photo to outlook
                //ToDo: add google photo to outlook with new Google API
                //Stream stream = _googleService.GetPhoto(match.GoogleContact);
                using (var googlePhoto = Utilities.GetGoogleContactPhoto(this, gc))
                {
                    if (googlePhoto != null)    // Google may have an invalid photo
                    {
                        oc.SetOutlookPhoto(googlePhoto);
                        ContactPropertiesUtils.SetOutlookGoogleId(oc, gc);
                        Save(ref oc);
                    }
                }
            }
            else
            {
                var hasOutlookPhoto = oc.HasPhoto();
                if (hasOutlookPhoto)
                {
                    oc.RemovePicture();
                    ContactPropertiesUtils.SetOutlookGoogleId(oc, gc);
                    Save(ref oc);
                }
            }
        }

        public ContactGroup SaveGoogleGroup(ContactGroup group)
        {
            var groupsResource = new ContactGroupsResource(GooglePeopleService);

            //check if this group was not yet inserted on google.
            if (string.IsNullOrEmpty(group.ResourceName)) //ToDo: Check, maybe also use >0
            {
                //insert group.
                //var feedUri = new Uri(GroupsQuery.CreateGroupsUri("default"));
                var contactGroupRequest = new CreateContactGroupRequest()
                {
                    ContactGroup = group
                };

                try
                {
                    return groupsResource.Create(contactGroupRequest).Execute();
                }
                catch (ProtocolViolationException)
                {
                    //TODO (obelix30)
                    //http://stackoverflow.com/questions/23804960/contactsrequest-insertfeeduri-newentry-sometimes-fails-with-system-net-protoc
                    try
                    {
                        return groupsResource.Create(contactGroupRequest).Execute();
                    }
                    catch (Exception ex)
                    {
                        Log.Debug(ex, $"ContactGroup dump: {group}");
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, $"ContactGroup dump: {group}");
                    throw;
                }
            }
            else
            {
                try
                {
                    var contactGroupRequest = new UpdateContactGroupRequest()
                    {
                        ContactGroup = group
                    };
                    //group already present in google. just update
                    return groupsResource.Update(contactGroupRequest, group.ResourceName).Execute();
                }
                catch
                {
                    //TODO: save google group xml for diagnistics
                    throw;
                }
            }
        }

        /// <summary>
        /// Updates Google contact from Outlook (including groups/categories)
        /// </summary>
        internal void UpdateContact(ContactItem master, Person slave)
        {
            ContactSync.UpdateContact(master, slave, UseFileAs);
            OverwriteContactGroups(master, slave);
        }

        /// <summary>
        /// Updates Google contact from Outlook (including groups/categories)
        /// </summary>
        public void UpdateContact(ContactItem master, Person slave, ContactMatch match)
        {
            match.GoogleContactDirty = true;
            //ContactSync.UpdateContact(master, slave, UseFileAs);
            UpdateContact(master, slave);
        }

        /// <summary>
        /// Updates Outlook contact from Google (including groups/categories)
        /// </summary>
        public void UpdateContact(Person master, ContactItem slave, bool googleContactDirty, bool matchedById)
        {
            ContactSync.UpdateContact(master, slave, UseFileAs);
            OverwriteContactGroups(master, slave);
            
            //If changed and not yet saved or matched by unqiue properties (not ID), assign syncId and save the contacts
            if (!slave.Saved || googleContactDirty || !matchedById)
            {// -- Immediately save the Outlook contact (including groups) so it can be released, and don't do it in the save loop later
                SaveOutlookContact(ref master, slave);
                SyncedCount++;
                Log.Information($"Updated contact from Google to Outlook: \"{slave.ToLogString()}\".");
            }
        }

        /// <summary>
        /// Updates Google contact's groups from Outlook contact
        /// </summary>
        private void OverwriteContactGroups(ContactItem master, Person slave)
        {
            //Contact group name "Starred in Android" is a reserved legacy name, was used in old Contact API. 
            //Backward compliancy by using the system "contactGroups/starred" group instead
            if (!string.IsNullOrEmpty(master.Categories) && master.Categories.Contains("Starred in Android"))
            {
                Utilities.RemoveOutlookGroup(master, "Starred in Android");
                Utilities.AddOutlookGroup(master, "starred");
            }

            var currentGroups = Utilities.GetGoogleGroups(this, slave);

            // get outlook categories
            var cats = Utilities.GetOutlookGroups(master.Categories);


            // remove obsolete groups
            var remove = new Collection<ContactGroup>();
            bool found;
            foreach (var group in currentGroups)
            {
                found = false;
                foreach (var cat in cats)
                {
                    if (group.Name == cat)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    remove.Add(group);
                }
            }
            while (remove.Count != 0)
            {
                Utilities.RemoveGoogleGroup(slave, remove[0]);
                remove.RemoveAt(0);
            }

            // add new groups
            ContactGroup g;
            foreach (var cat in cats)
            {
                if (!Utilities.ContainsGroup(this, slave, cat))
                {
                    // add group to contact
                    g = GetGoogleGroupByName(cat);
                    if (g == null)
                    {
                        // try to create group again (if not yet created before
                        g = CreateGroup(cat);

                        if (g != null)
                        {
                            g = SaveGoogleGroup(g);
                            if (g != null)
                            {
                                GoogleGroups.Add(g);
                            }
                            else
                            {
                                Log.Warning($"Google Groups were supposed to be created prior to saving a contact. Unfortunately the group '{cat}' couldn't be saved on Google side and was not assigned to the contact: {master.ToLogString()}");
                            }
                        }
                        else
                        {
                            Log.Warning($"Google Groups were supposed to be created prior to saving a contact. Unfortunately the group '{cat}' couldn't be created and was not assigned to the contact: {master.ToLogString()}");
                        }
                    }

                    if (g != null)
                    {
                        Utilities.AddGoogleGroup(slave, g);
                    }
                }
            }

            //add system ContactGroup My Contacts            
            /*if (!Utilities.ContainsGroup(this, slave, myContactsGroup))
            {
                // add group to contact
                g = GetGoogleGroupByName(myContactsGroup);
                if (g == null)
                {
                    throw new Exception($"Google {myContactsGroup} doesn't exist");
                }
                Utilities.AddGoogleGroup(slave, g);
            }*/
        }

        /// <summary>
        /// Updates Outlook contact's categories (groups) from Google groups
        /// </summary>
        private void OverwriteContactGroups(Person master, ContactItem slave)
        {
            var newGroups = Utilities.GetGoogleGroups(this, master);

            var newCats = new List<string>(newGroups.Count);
            foreach (var group in newGroups)
            {   //Not needed anymore with Google People API: Only add groups that are no SystemGroup (e.g. "System ContactGroup: Meine Kontakte") automatically tracked by Google
                if (group.Name != null) // && !group.Name.Equals(myContactsGroup))
                {
                    newCats.Add(group.Name);
                }
            }

            slave.Categories = string.Join(", ", newCats.ToArray());
        }

        /// <summary>
        /// Resets associantions of Outlook contacts with Google contacts via user props
        /// and resets associantions of Google contacts with Outlook contacts via extended properties.
        /// </summary>
        public void ResetContactMatches()
        {
            Debug.Assert(OutlookContacts != null, "Outlook Contacts object is null - this should not happen. Please inform Developers.");
            Debug.Assert(GoogleContacts != null, "Google Contacts object is null - this should not happen. Please inform Developers.");

            try
            {
                if (string.IsNullOrEmpty(SyncProfile))
                {
                    Log.Error("Must set a sync profile. This should be different on each user/computer you sync on.");
                    return;
                }

                lock (_syncRoot)
                {
                    Log.Information("Resetting Google Person matches...");
                    foreach (var gc in GoogleContacts)
                    {
                        try
                        {
                            if (gc != null)
                            {
                                ResetMatch(gc);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Warning($"The match of Google contact {ContactMatch.GetName(gc)} couldn't be reset: {ex.Message}");
                        }
                    }

                    Log.Information("Resetting Outlook Contact matches...");

                    var item = OutlookContacts.GetFirst();
                    while (item != null)
                    {
                        //"is" operator creates an implicit variable (COM leak), so unfortunately we need to avoid pattern matching
#pragma warning disable IDE0019 // Use pattern matching
                        var oc = item as ContactItem;
#pragma warning restore IDE0019 // Use pattern matching

                        if (oc != null)
                        {
                            try
                            {
                                ResetMatch(oc);
                            }
                            catch (Exception ex)
                            {
                                var name = oc.ToLogString();
                                if (string.IsNullOrWhiteSpace(name))
                                {
                                    Log.Warning($"The match of Outlook contact couldn't be reset: {ex.Message}");
                                }
                                else
                                {
                                    Log.Warning($"The match of Outlook contact {name} couldn't be reset: {ex.Message}");
                                }
                            }
                        }
                        else
                        {
                            Log.Debug("Empty Outlook contact found (maybe distribution list). Skipping");
                        }
                        Marshal.ReleaseComObject(item);
                        item = OutlookContacts.GetNext();
                    }
                }
            }
            finally
            {
                GoogleContacts = null;
            }
        }

        /// <summary>
        /// Resets associations of Outlook appointments with Google appointments via user props
        /// and vice versa
        /// </summary>
        public void ResetOutlookAppointmentMatches(bool deleteOutlookAppointments)
        {
            Debug.Assert(OutlookAppointments != null, "Outlook Appointments object is null - this should not happen. Please inform Developers.");

            lock (_syncRoot)
            {
                Log.Information("Resetting Outlook appointment matches...");
                //1 based array
                for (var i = OutlookAppointments.Count; i >= 1; i--)
                {
                    AppointmentItem oa = null;

                    try
                    {
                        oa = OutlookAppointments[i] as AppointmentItem;
                        if (oa == null)
                        {
                            Log.Warning("Empty Outlook appointment found. Skipping");
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        //this is needed because some appointments throw exceptions
                        Log.Warning($"Accessing Outlook appointment threw an exception. Skipping: {ex.Message}");
                        continue;
                    }

                    if (deleteOutlookAppointments)
                    {
                        oa.Delete();
                    }
                    else
                    {
                        try
                        {
                            ResetMatch(oa);
                        }
                        catch (Exception ex)
                        {
                            Log.Warning($"The match of Outlook appointment {oa.ToLogString()} couldn't be reset: {ex.Message}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Reset the match link between Google and Outlook contact        
        /// </summary>
        public Person ResetMatch(Person gc)
        {
            if (gc != null)
            {
                ContactPropertiesUtils.ResetGoogleOutlookId(gc);
                return SaveGoogleContact(gc);
            }
            else
            {
                return gc;
            }
        }

        /// <summary>
        /// Reset the match link between Outlook and Google contact
        /// </summary>
        public void ResetMatch(ContactItem oc)
        {
            if (oc != null)
            {
                ContactPropertiesUtils.ResetOutlookGoogleId(this, oc);
                Save(ref oc);
            }
        }

        /// <summary>
        /// Reset the match link between Outlook and Google appointment
        /// </summary>
        public void ResetMatch(AppointmentItem oa)
        {
            if (oa != null)
            {
                AppointmentPropertiesUtils.ResetOutlookGoogleId(oa);
                Save(ref oa);
            }
        }

        public ContactMatch ContactByProperty(string name, string email)
        {
            foreach (var m in Contacts)
            {
                var fileAs = ContactPropertiesUtils.GetGoogleFileAsValue(m.GoogleContact);
                var primaryEmail = ContactPropertiesUtils.GetGooglePrimaryEmailValue(m.GoogleContact);
                var googleName = ContactPropertiesUtils.GetGoogleUnstructuredName(m.GoogleContact);
                if (!string.IsNullOrEmpty(primaryEmail) && primaryEmail.Equals(email, StringComparison.InvariantCultureIgnoreCase) ||
                    !string.IsNullOrEmpty(fileAs) && fileAs.Equals(name, StringComparison.InvariantCultureIgnoreCase) ||
                    !string.IsNullOrEmpty(googleName) && googleName.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                {
                    return m;
                }
                else if (m.OutlookContact != null && (
                    (m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email) ||
                    m.OutlookContact.FullName == name ||
                    m.OutlookContact.FileAs == name))
                {
                    return m;
                }
            }
            return null;
        }

        /// <summary>
        /// Used to find duplicates.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public Collection<OutlookContactInfo> OutlookContactByProperty(string name, string value)
        {
            var col = new Collection<OutlookContactInfo>();

            try
            {
                var item = OutlookContacts.Find($"[{name}] = \"{value}\"") as ContactItem;
                while (item != null)
                {
                    col.Add(new OutlookContactInfo(item, this));
                    item = OutlookContacts.FindNext() as ContactItem;
                }
            }
            catch (Exception)
            {
                //TODO: should not get here.
            }

            return col;
        }

        public ContactGroup GetGoogleGroupByResourceName(string resourceName)
        {
            //return GoogleGroups.FindById(new string(id)) as ContactGroup;
            foreach (var group in GoogleGroups)
            {
                if (resourceName == group.ResourceName) //ToDo: Check
                {
                    return group;
                }
            }
            return null;
        }

        public ContactGroup GetGoogleGroupByName(string name)
        {
            foreach (var group in GoogleGroups)
            {
                if (group.Name == name)
                {
                    return group;
                }
            }
            return null;
        }


        public Person GetGoogleContact(string gid)
        {
            var slash = gid.LastIndexOf("/"); //ToDo: For Backward compatibility with old GoogleContact-API: remove the prefix from the id (e.g. 1c0d39680d700698 from http://www.google.com/m8/feeds/contacts/saller.flo%40gmail.com/base/1c0d39680d700698)
            gid = gid.Substring(slash + 1);

            var gc = GetGoogleContactById(gid);

            if (gc != null)
            {
                return gc;
            }
            else
            {
                return LoadGoogleContacts(gid);
            }
        }

        public Person GetGoogleContactById(string id)
        {
            if (!string.IsNullOrEmpty(id))
            {
                foreach (var gc in GoogleContacts)
                {
                    //var slash = id.LastIndexOf("/"); //ToDo: For Backward compatibility with old GoogleContact-API: remove the prefix from the id (e.g. 1c0d39680d700698 from http://www.google.com/m8/feeds/contacts/saller.flo%40gmail.com/base/1c0d39680d700698)
                    //id = id.Substring(slash+1);

                    var gid = ContactPropertiesUtils.GetGoogleId(gc);
                    if (gid.Equals(id, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return gc;
                    }
                }
            }
            return null;
        }

        public Event GetGoogleAppointmentById(string id)
        {
            foreach (var ga in GoogleAppointments)
            {
                if (ga.Id.Equals(id))
                {
                    return ga;
                }
            }

            if (AllGoogleAppointments != null)
            {
                foreach (var ga in AllGoogleAppointments)
                {
                    if (ga.Id.Equals(id))
                    {
                        return ga;
                    }
                }
            }

            return null;
        }

        public static AppointmentItem GetOutlookAppointmentById(string id)
        {
            var o = OutlookNameSpace.GetItemFromID(id);
            //"is" operator creates an implicit variable (COM leak), so unfortunately we need to avoid pattern matching
#pragma warning disable IDE0019 // Use pattern matching
            var oa = o as AppointmentItem;
#pragma warning restore IDE0019 // Use pattern matching

            return oa;
        }

        public ContactItem GetOutlookContactById(string id)
        {
            if (!string.IsNullOrEmpty(id))
            {
                for (var i = OutlookContacts.Count; i >= 1; i--)
                {
                    ContactItem oc;
                    try
                    {
                        oc = OutlookContacts[i] as ContactItem;
                        if (oc == null)
                        {
                            continue;
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    if (id.Equals(ContactPropertiesUtils.GetOutlookId(oc), StringComparison.InvariantCultureIgnoreCase))
                    {
                        return oc;
                    }
                }
            }
            return null;
        }

        public ContactGroup CreateGroup(string name)
        {
            var group = new ContactGroup
            {
                Name = name
            };
            //group.GroupEntry.Dirty = true;
            return group;
        }

        public static bool AreEqual(ContactItem oc1, ContactItem oc2)
        {
            return oc1.Email1Address == oc2.Email1Address;
        }

        public static int IndexOf(Collection<ContactItem> col, ContactItem oc)
        {
            for (var i = 0; i < col.Count; i++)
            {
                if (AreEqual(col[i], oc))
                {
                    return i;
                }
            }
            return -1;
        }

        internal void DebugContacts()
        {
            var oCount = string.Empty;
            var gCount = string.Empty;
            var mCount = string.Empty;

            if (SyncContacts)
            {
                oCount = $"Outlook Contact Count: {OutlookContacts.Count}";
                gCount = $"Google Person Count: {GoogleContacts.Count}";
                mCount = $"Matches Count: {Contacts.Count}";
            }

            if (SyncAppointments)
            {
                oCount = $"Outlook appointments Count: {OutlookAppointments.Count}";
                gCount = $"Google appointments Count: {GoogleAppointments.Count}";
                mCount = $"Matches Count: {Appointments.Count}";
            }

            MessageBox.Show($"DEBUG INFORMATION\nPlease submit to developer:\n\n{oCount}\n{gCount}\n{mCount}", "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        public static ContactItem CreateOutlookContactItem(string syncContactsFolder)
        {
            Outlook.ContactItem outlookContact = null;
            MAPIFolder contactsFolder = null;
            Items items = null;
            try
            {
                contactsFolder = OutlookNameSpace.GetFolderFromID(syncContactsFolder);
                items = contactsFolder.Items;
                outlookContact = items.Add(OlItemType.olContactItem) as ContactItem;
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
                if (contactsFolder != null)
                {
                    Marshal.ReleaseComObject(contactsFolder);
                }
            }

            return outlookContact;
        }

        public static AppointmentItem CreateOutlookAppointmentItem(string syncAppointmentsFolder)
        {
            Outlook.AppointmentItem outlookAppointment = null;
            MAPIFolder appointmentsFolder = null;
            Items items = null;
            try
            {
                appointmentsFolder = OutlookNameSpace.GetFolderFromID(syncAppointmentsFolder);
                items = appointmentsFolder.Items;
                outlookAppointment = items.Add(OlItemType.olAppointmentItem) as AppointmentItem;
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
                if (appointmentsFolder != null)
                {
                    Marshal.ReleaseComObject(appointmentsFolder);
                }
            }

            return outlookAppointment;
        }

        public void Dispose()
        {
            if (GoogleCalendarService != null)
            {
                ((IDisposable)GoogleCalendarService).Dispose();
            }
        }
    }

    public enum SyncOption
    {
        MergePrompt,
        MergeOutlookWins,
        MergeGoogleWins,
        OutlookToGoogleOnly,
        GoogleToOutlookOnly,
    }
}