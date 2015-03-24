using Microsoft.Win32;
using Segment;
using System;
using System.Diagnostics;
using System.Net;
using System.Text;

namespace Setup
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            // segment.io write key
            const string WRITE_KEY = "bjgcaeb4bj";

            // initialize segment.io, if debugging, do not use async
#if DEBUG
            Analytics.Initialize(WRITE_KEY, new Config().SetAsync(false));
            System.Diagnostics.Debug.Print("Debugging, calls to segment.io are synchronous");
#else
            Analytics.Initialize(WRITE_KEY);
#endif

            writeLine("Installing PowerPoint Warrior Toolbar");
            writeLine("If you have any questions, or if installing the product is not successful, please contact support at info@ppwarrior.com");
            writeLine("");

            // create segment properties to track versions etc.
            var props = new Segment.Model.Properties();
            props.Add("Windows version", getWindowsVersion());
            props.Add("VSTO version", getVstoVersion());
            props.Add("Office version", getOfficeVersion());
            string netVersions = getNetVersionsAll();
            props.Add(".NET version all", netVersions);
            props.Add(".NET version latest", getNetVersion(netVersions));

            Analytics.Client.Track(Environment.MachineName + "\\" + Environment.UserName, "Started installer", props);

            string appdata = System.Windows.Forms.Application.UserAppDataPath;
            string location = appdata + "\\warrior.exe";

            writeLine("Downloading setup...");

            WebClient client = new WebClient();
            client.DownloadFile(@"https://ppwarrior.blob.core.windows.net/install/setup.exe", location);

            writeLine("Starting setup...");

            Process.Start(location);

            writeLine("Setup started, quitting...");

            System.Threading.Thread.Sleep(3000);

            Environment.Exit(1);
        }

        private static string getNetVersion(string all)
        {
            if (all.IndexOf("4.5") > 0)
                return "4.5";
            if (all.IndexOf("4.0") > 0)
                return "4.0";
            if (all.IndexOf("3.5") > 0)
                return "3.5";
            if (all.IndexOf("2.0") > 0)
                return "2.0";
            return "Unknown";
        }

        private static string getOfficeVersion()
        {
            // http://xltoolbox.sourceforge.net/blog/2015/01/net-vsto-add-ins-getting-prerequisites-right/

            if (regKeyExists(@"SOFTWARE\Microsoft\Office\15.0"))
                return "15.0";
            if (regKeyExists(@"SOFTWARE\Microsoft\Office\14.0"))
                return "14.0";
            if (regKeyExists(@"SOFTWARE\Microsoft\Office\12.0"))
                return "12.0";
            return "Unknown";
        }

        private static bool regKeyExists(string v)
        {
            if (RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, "").OpenSubKey(v) != null)
                return true;
            else
                return false;
        }

        private static string getVstoVersion()
        {
            if (regKeyExists(@"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R") ||
                regKeyExists(@"SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R"))
                return "4.0";
            else
                return "Unknown";
        }

        private static string getWindowsVersion()
        {
            return string.Format("{0}.{1}", Environment.OSVersion.Version.Major, Environment.OSVersion.Version.Minor);
        }

        private static void writeLine(string v)
        {
            Console.WriteLine(v);
        }

        // https://msdn.microsoft.com/en-us/library/hh925568%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396#net_c
        private static string getNetVersionsAll()
        {
            StringBuilder sb = new StringBuilder();

            // Opens the registry key for the .NET Framework entry. 
            using (RegistryKey ndpKey =
                RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, "").
                OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\"))
            {
                foreach (string versionKeyName in ndpKey.GetSubKeyNames())
                {
                    if (versionKeyName.StartsWith("v"))
                    {

                        RegistryKey versionKey = ndpKey.OpenSubKey(versionKeyName);
                        string name = (string)versionKey.GetValue("Version", "");
                        string sp = versionKey.GetValue("SP", "").ToString();
                        string install = versionKey.GetValue("Install", "").ToString();
                        if (install == "") //no install info, must be later.
                            sb.AppendLine(versionKeyName + "  " + name);
                        else
                        {
                            if (sp != "" && install == "1")
                            {
                                sb.AppendLine(versionKeyName + "  " + name + "  SP" + sp);
                            }

                        }
                        if (name != "")
                        {
                            continue;
                        }
                        foreach (string subKeyName in versionKey.GetSubKeyNames())
                        {
                            RegistryKey subKey = versionKey.OpenSubKey(subKeyName);
                            name = (string)subKey.GetValue("Version", "");
                            if (name != "")
                                sp = subKey.GetValue("SP", "").ToString();
                            install = subKey.GetValue("Install", "").ToString();
                            if (install == "") //no install info, must be later.
                                sb.AppendLine(versionKeyName + "  " + name);
                            else
                            {
                                if (sp != "" && install == "1")
                                {
                                    sb.AppendLine("  " + subKeyName + "  " + name + "  SP" + sp);
                                }
                                else if (install == "1")
                                {
                                    sb.AppendLine("  " + subKeyName + "  " + name);
                                }

                            }

                        }

                    }
                }
            }

            return sb.ToString();
        }

    }
}
