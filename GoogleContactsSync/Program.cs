using Serilog;
using Serilog.Events;
using Serilog.Exceptions;
using System;
using System.Globalization;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using static GoContactSyncMod.InMemorySink;

namespace GoContactSyncMod
{
    internal static class Program
    {
        //"ACBBBC09-F76C-4874-AAFF-4F3353A5A5A6"
        private static readonly string MUTEXGUID = (Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(GuidAttribute), false)[0] as GuidAttribute).Value;
        private static Mutex mutex;
        private static readonly InMemorySink sink = new InMemorySink();

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");

            //for very rare setups we could have error:
            //"The client and server cannot communicate, because they do not possess a common algorithm"
            //see also https://github.com/googleapis/google-api-dotnet-client/issues/911
            //so lets add manually option to use TLS 1.2
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            ConfigureLogger();

            //prevent more than one instance of the program
            if (IsRunning())
            {   //Instance already exists, so show only Main-Window  
                NativeMethods.PostMessage((IntPtr)NativeMethods.HWND_BROADCAST, NativeMethods.WM_GCSM_SHOWME, IntPtr.Zero, IntPtr.Zero);
                return;
            }
            else
            {
                RegisterEventHandlers();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(SettingsForm.Instance);
            }
            GC.KeepAlive(mutex);
        }

        public static void EnableLogHandler(LogHandler log)
        {
            sink.OnLogReceived += log;
        }

        public static void DisableLogHandler(LogHandler log)
        {
            sink.OnLogReceived -= log;
        }

        private static void ConfigureLogger()
        {
            var Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\GoContactSyncMOD\\";

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .Enrich.WithExceptionDetails()
                .WriteTo.Sink(sink, restrictedToMinimumLevel: LogEventLevel.Information)
                .WriteTo.File(Folder + "log.txt",
                    rollingInterval: RollingInterval.Day,
                    rollOnFileSizeLimit: true,
                    fileSizeLimitBytes: 10 * 1024 * 1024,  //roll file after 10MB
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3} {Message:lj}{NewLine}{Exception}")
                .CreateLogger();
        }

        private static void RegisterEventHandlers()
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
        }

        public static bool IsRunning()
        {
            mutex = new Mutex(true, MUTEXGUID, out _);
            return !mutex.WaitOne(TimeSpan.Zero, true);
        }

        /// <summary>
        /// Fallback. If there is some try/catch missing we will handle it here, just before the application quits unhandled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception exception)
            {
                ErrorHandler.Handle(exception);
            }
        }
    }
}