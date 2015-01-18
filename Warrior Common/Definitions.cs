using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Deployment.Application;
using System.Net;
using System.Text.RegularExpressions;

namespace WarriorCommon
{
    public static class Information
    {
        internal static string GetUserName()
        {
            return Environment.MachineName + "\\" + Environment.UserName;
        }

        internal static string GetAssemblyVersion()
        {
            return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        internal static string GetClickOnceVersion()
        {
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            return "Debug";
        }

        internal static string GetExternalIp()
        {
            string externalIP;
            externalIP = (new WebClient()).DownloadString("http://checkip.dyndns.org/");
            externalIP = (new Regex(@"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"))
                         .Matches(externalIP)[0].ToString();
            return externalIP;
        }
    }
}
