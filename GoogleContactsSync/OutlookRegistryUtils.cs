using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;

namespace GoContactSyncMod
{
    internal static class OutlookRegistryUtils
    {
        public static string GetOutlookVersion()
        {
            var ret = string.Empty;
            RegistryKey registryOutlookKey = null;

            try
            {
                var outlookVersion = GetMajorVersion(GetOutlookPath());
                ret = GetMajorVersionToString(outlookVersion);

                var outlookKey = @"Software\Wow6432Node\Microsoft\Office\" + outlookVersion + @".0\Outlook";
                registryOutlookKey = Registry.LocalMachine.OpenSubKey(outlookKey, false);
                if (registryOutlookKey == null)
                {
                    outlookKey = @"Software\Microsoft\Office\" + outlookVersion + @".0\Outlook";
                    registryOutlookKey = Registry.LocalMachine.OpenSubKey(outlookKey, false);
                }
                if (registryOutlookKey != null)
                {
                    var bitness = registryOutlookKey.GetValue(@"Bitness", string.Empty).ToString();
                    if (string.IsNullOrEmpty(bitness))
                    {
                        return ret;
                    }
                    else
                    {
                        if (bitness == @"x86")
                        {
                            return ret + @" (32-bit)";
                        }
                        else
                        {
                            return bitness == @"x64" ? ret + @" (64-bit)" : ret;
                        }
                    }
                }
            }
            catch
            {
            }
            finally
            {
                if (registryOutlookKey != null)
                {
                    registryOutlookKey.Close();
                }
            }

            return ret;
        }

        public static string GetPossibleErrorDiagnosis()
        {
            var outlookVersion = GetMajorVersion(GetOutlookPath());
            var diagnosis = CheckOfficeRegistry(outlookVersion);
            var c2r = GetClickToRunVersion();
            var ret = "Could not connect to 'Microsoft Outlook'.\r\nYou have " + GetMajorVersionToString(outlookVersion) + " installed.";
            if (!string.IsNullOrEmpty(c2r))
            {
                ret += "Click to Run was also detected (" + c2r + ").";
            }

            return ret + "\r\n" + diagnosis;
        }

        private static string CheckOfficeRegistry(int outlookVersion)
        {
            var registryVersion = ConvertMajorVersionToRegistryVersion(outlookVersion);

            var interfaceVersion = @"Interface\{00063001-0000-0000-C000-000000000046}\TypeLib";
            var interfaceKey = Registry.ClassesRoot.OpenSubKey(interfaceVersion, false);

            if (interfaceKey == null)
            {
                interfaceVersion = @"WOW6432Node\Interface\{00063001-0000-0000-C000-000000000046}\TypeLib";
                interfaceKey = Registry.ClassesRoot.OpenSubKey(interfaceVersion, false);
            }

            if (interfaceKey != null)
            {
                var typeLib = interfaceKey.GetValue(string.Empty).ToString();
                if (typeLib != "{00062FFF-0000-0000-C000-000000000046}")
                {
                    return "Your registry " + interfaceKey.ToString() + " points to TypeLib " + typeLib + " and should to {00062FFF-0000-0000-C000-000000000046}" + registryVersion + ".\r\nPlease read FAQ (https://sourceforge.net/p/googlesyncmod/faq/2016/05/fixing-office-installation/) and fix your Office installation";
                }

                var versionObj = interfaceKey.GetValue("Version");
                if (versionObj != null)
                {
                    var version = versionObj.ToString();
                    if (version != registryVersion)
                    {
                        return "Your registry " + interfaceKey.ToString() + " points to version " + version + " and your Outlook is installed with version " + registryVersion + ".\r\nPlease read FAQ (https://sourceforge.net/p/googlesyncmod/faq/2016/05/fixing-office-installation/) and fix your Office installation";
                    }
                }
                else
                {
                    return "There is no version key in registry " + interfaceKey.ToString() + ".\r\nPlease read FAQ (https://sourceforge.net/p/googlesyncmod/faq/2016/05/fixing-office-installation/) and fix your Office installation";
                }
            }
            else
            {
                return "Cannot open registry " + interfaceVersion + ".\r\nPlease read FAQ (https://sourceforge.net/p/googlesyncmod/faq/2016/05/fixing-office-installation/) and fix your Office installation";
            }

            if (!string.IsNullOrEmpty(registryVersion))
            {
                var RegKey = @"TypeLib\{00062FFF-0000-0000-C000-000000000046}\" + registryVersion + @"\0\";

                var mainKey = Registry.ClassesRoot.OpenSubKey(RegKey + "win32", false);
                if (mainKey != null)
                {
                    var path = mainKey.GetValue(string.Empty).ToString();
                    if (!File.Exists(path))
                    {
                        return "Your registry " + mainKey.ToString() + " points to file " + path + " and this file does not exist.\r\nPlease read FAQ (https://sourceforge.net/p/googlesyncmod/faq/2016/05/fixing-office-installation/) and fix your Office installation";
                    }
                }
                mainKey = Registry.ClassesRoot.OpenSubKey(RegKey + "win64", false);
                if (mainKey != null)
                {
                    var path = mainKey.GetValue(string.Empty).ToString();
                    if (!File.Exists(path))
                    {
                        return "Your registry " + mainKey.ToString() + " points to file " + path + " and this file does not exist.\r\nPlease read FAQ (https://sourceforge.net/p/googlesyncmod/faq/2016/05/fixing-office-installation/) and fix your Office installation";
                    }
                }

                mainKey = Registry.ClassesRoot.OpenSubKey(@"TypeLib\{00062FFF-0000-0000-C000-000000000046}\", false);
                if (mainKey != null)
                {
                    var keys = mainKey.GetSubKeyNames();
                    if (keys.Length > 1)
                    {
                        var allKeys = "";
                        for (var i = 0; i < keys.Length; i++)
                        {
                            var element = keys[i];
                            if (element != registryVersion)
                            {
                                allKeys = allKeys + '"' + ConvertRegistryVersionToString(element) + '"' + ",";
                            }
                        }
                        allKeys = allKeys.Substring(0, allKeys.Length - 1);
                        return "Your registry " + mainKey.ToString() + " points to Office versions: " + allKeys + " other than you have installed \"" + ConvertRegistryVersionToString(registryVersion) + "\".\r\nPlease read FAQ (https://sourceforge.net/p/googlesyncmod/faq/2016/05/fixing-office-installation/) and fix your Office installation";
                    }
                }
            }
            return "Please read FAQ (https://sourceforge.net/p/googlesyncmod/faq/2016/05/fixing-office-installation/) and fix your Office installation";
        }

        private static string GetClickToRunVersion()
        {
            string[] keys = { @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\", @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\" };
            string[] versions = { "9.0", "10.0", "11.0", "12.0", "14.0", "15.0", "16.0" };

            for (var j = 0; j < keys.Length; j++)
            {
                for (var i = 0; i < versions.Length; i++)
                {
                    try
                    {
                        var registryKey = Registry.LocalMachine.OpenSubKey(keys[j] + versions[i] + @"\Common\InstallRoot\Virtual\VirtualOutlook", false);
                        if (registryKey != null)
                        {
                            return versions[i];
                        }
                    }
                    finally
                    {
                    }
                }
            }

            return string.Empty;
        }

        private static string GetOutlookPath()
        {
            const string regKey = @"Software\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe";
            var toReturn = string.Empty;

            try
            {
                var mainKey = Registry.CurrentUser.OpenSubKey(regKey, false);
                if (mainKey != null)
                {
                    toReturn = mainKey.GetValue(string.Empty).ToString();
                }
            }
            catch
            { }

            if (string.IsNullOrEmpty(toReturn))
            {
                try
                {
                    var mainKey = Registry.LocalMachine.OpenSubKey(regKey, false);
                    if (mainKey != null)
                    {
                        toReturn = mainKey.GetValue(string.Empty).ToString();
                    }
                }
                catch
                { }
            }

            return toReturn;
        }

        private static int GetMajorVersion(string path)
        {
            var toReturn = 0;
            if (File.Exists(path))
            {
                try
                {
                    toReturn = FileVersionInfo.GetVersionInfo(path).FileMajorPart;
                }
                catch
                { }
            }
            return toReturn;
        }

        private static string ConvertMajorVersionToRegistryVersion(int version)
        {
            switch (version)
            {
                case 9: return "9.0";
                case 10: return "9.1";
                case 11: return "9.2";
                case 12: return "9.3";
                case 14: return "9.4";
                case 15: return "9.5";
                case 16: return "9.6";
                default: return Convert.ToString(version);
            }
        }

        private static string ConvertRegistryVersionToString(string version)
        {
            switch (version)
            {
                case "9.0": return "Office 2000";
                case "9.1": return "Office XP";
                case "9.2": return "Office 2003";
                case "9.3": return "Office 2007";
                case "9.4": return "Office 2010";
                case "9.5": return "Office 2013";
                case "9.6": return "Office 2016";
                default: return version;
            }
        }

        private static string GetMajorVersionToString(int version)
        {
            switch (version)
            {
                case 7: return "Office 97";
                case 8: return "Office 98";
                case 9: return "Office 2000";
                case 10: return "Office XP";
                case 11: return "Office 2003";
                case 12: return "Office 2007";
                case 14: return "Office 2010";
                case 15: return "Office 2013";
                case 16:
                    {
                        var regKey = @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun";
                        try
                        {
                            var mainKey = Registry.LocalMachine.OpenSubKey(regKey, false);
                            if (mainKey == null)
                            {
                                regKey = @"Software\Microsoft\Office\ClickToRun";
                                mainKey = Registry.LocalMachine.OpenSubKey(regKey, false);
                                if (mainKey == null)
                                {
                                    return "Office 2016 / Office 2019 / Office 365";
                                }
                            }
                        }
                        catch
                        {
                            return "Office 2016 / Office 2019 / Office 365";
                        }
                        //regKey = @"Software\Microsoft\Office\ClickToRun\Configuration\ProductReleaseIDs";
                        regKey = @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration\ProductReleaseIDs";
                        try
                        {
                            var mainKey = Registry.LocalMachine.OpenSubKey(regKey, false);
                            if (mainKey != null)
                            {
                                var s = mainKey.GetValue(string.Empty).ToString();
                                if (s == "ProPlus2019Retail")
                                {
                                    return "Office 2019";
                                }
                                else
                                {
                                    return s == "O365ProPlusRetail" ? "Office 365" : s;
                                }
                            }
                            else
                            {
                                regKey = @"Software\Microsoft\Office\ClickToRun\Configuration\ProductReleaseIDs";
                                mainKey = Registry.LocalMachine.OpenSubKey(regKey, false);
                                if (mainKey != null)
                                {
                                    var s = mainKey.GetValue(string.Empty).ToString();
                                    if (s == "ProPlus2019Retail")
                                    {
                                        return "Office 2019";
                                    }
                                    else
                                    {
                                        return s == "O365ProPlusRetail" ? "Office 365" : s;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            return "Office 2016 / Office 2019 / Office 365";
                        }
                        return "Office 2016 / Office 2019 / Office 365";
                    }
                default: return "unknown Office version (" + version + ")";
            }
        }
    }
}
