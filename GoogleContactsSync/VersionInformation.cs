using Serilog;
using System;
using System.Diagnostics;
using System.Management;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal static class VersionInformation
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        public enum OutlookMainVersion
        {
            Outlook2002,
            Outlook2003,
            Outlook2007,
            Outlook2010,
            Outlook2013,
            Outlook_2016_or_2019_or_365,
            OutlookUnknownVersion,
            OutlookNoInstance
        }

        public static OutlookMainVersion GetOutlookVersion(Outlook.Application appVersion)
        {
            try
            {
                if (appVersion == null)
                {
                    appVersion = new Outlook.Application();
                }

                switch (appVersion.Version.ToString().Substring(0, 2))
                {
                    case "10":
                        return OutlookMainVersion.Outlook2002;
                    case "11":
                        return OutlookMainVersion.Outlook2003;
                    case "12":
                        return OutlookMainVersion.Outlook2007;
                    case "14":
                        return OutlookMainVersion.Outlook2010;
                    case "15":
                        return OutlookMainVersion.Outlook2013;
                    case "16":
                        return OutlookMainVersion.Outlook_2016_or_2019_or_365;
                    default:
                        {
                            Log.Debug("Unknown Outlook Version " + appVersion.Version.ToString().Substring(0, 2));
                            return OutlookMainVersion.OutlookUnknownVersion;
                        }
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                return OutlookMainVersion.OutlookUnknownVersion;
            }
        }

        /// <summary>
        /// detect windows main version
        /// </summary>
        public static string GetWindowsVersion()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("root\\CIMV2",
                        "SELECT Caption, OSArchitecture, Version FROM Win32_OperatingSystem"))
                {
                    foreach (ManagementObject managementObject in searcher.Get())
                    {
                        var versionString = managementObject["Caption"].ToString() + " (" +
                                               managementObject["OSArchitecture"].ToString() + "; " +
                                               managementObject["Version"].ToString() + ")";
                        return versionString;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
            }

            return "Unknown Windows Version";
        }

        public static Version GetGCSMVersion()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            var assemblyVersionNumber = new Version(fvi.FileVersion);

            return assemblyVersionNumber;
        }

        public static string GetGCSMDisplayVersion()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var info = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
                if (info != null && !string.IsNullOrWhiteSpace(info.InformationalVersion))
                {
                    return info.InformationalVersion;
                }

                var fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                return fvi.FileVersion;
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Exception");
                return GetGCSMVersion()?.ToString() ?? "unknown";
            }
        }

        public static async Task<bool> IsNewVersionAvailable(CancellationToken cancellationToken)
        {
            Log.Information("Reading version number from sf.net...");
            try
            {
                //specify to use TLS 1.2 as default connection
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var response = await _httpClient.GetAsync("https://sourceforge.net/projects/googlesyncmod/files/updates_v1.xml", HttpCompletionOption.ResponseHeadersRead, cancellationToken);

                response.EnsureSuccessStatusCode();
                var stream = await response.Content.ReadAsStreamAsync();
                var doc = XDocument.Load(stream);

                var strVersion = doc.Element("Version").Value;
                if (!string.IsNullOrEmpty(strVersion))
                {
                    var webVersionNumber = new Version(strVersion);
                    var localVersion = GetGCSMVersion();
                    var localDisplayVersion = GetGCSMDisplayVersion();
                    string addOn = $" (Installed: {((localVersion==null)?"null":localVersion.ToString())} / Available: {strVersion})";                    
                    //compare both versions
                    var result = webVersionNumber.CompareTo(localVersion);
                    if (result > 0)
                    {   //newer version found
                        Log.Information($"New version of GCSM available on sf.net{addOn}!");
                        return true;
                    }
                    else if(result < 0)
                    {
                        Log.Information($"You are using a PRE-RELEASE version of GCSM{addOn}.");
                        return false;
                    }
                    else
                    {   //older or same version found
                        Log.Information($"You are using the latest version of GCSM ({strVersion}). Installed build: {localDisplayVersion}.");
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log.Information("Could not read version number from sf.net...");
                Log.Debug(ex, "Exception");
                return false;
            }
        }
    }
}
