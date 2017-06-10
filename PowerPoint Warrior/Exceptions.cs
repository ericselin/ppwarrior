using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPoint_Warrior
{
	public static class Exceptions
	{
		public static void Handle(Exception ex, string officeVersion, string userEmail, bool showMessage = true)
		{
			Handle(ex, showMessage);
		}
		public static void Handle(Exception ex, bool showMessage = true)
		{
#if !DEBUG
			var userEmail = Properties.Settings.Default.UserEmail;
			// create raygun client and set appropriate flags
			RaygunClient _client = new RaygunClient("z2DRwwFDwxCn/6WrH/Irqg==");
			// raygun identifier is user name
			_client.User = Information.GetUserName();
			// insert email if we have it
			if (!string.IsNullOrEmpty(userEmail))
			{
				RaygunIdentifierMessage userInfo = new RaygunIdentifierMessage(_client.User);
				userInfo.Email = userEmail;
				_client.UserInfo = userInfo;
			}
			// set version to clickonce version
			_client.ApplicationVersion = Information.GetClickOnceVersion();
			// make custom data object
			var customData = new Dictionary<string, string>();
			try
			{
				customData.Add("ClickOnce Version", Information.GetClickOnceVersion());
				customData.Add("Office Version", Globals.ThisAddIn.Application.Version);
				customData.Add("Assembly Version", Information.GetAssemblyVersion());
				customData.Add("Current Selection", Globals.ThisAddIn.Application.ActiveWindow.Selection.Type.ToString());
				customData.Add("Active view pane", Globals.ThisAddIn.Application.ActiveWindow.ActivePane.ViewType.ToString());
			}
			catch (Exception iex)
			{
				ex = new Exception("Could not create custom data for exception", iex);
			}
			// send exception to raygun
			_client.Send(ex, null, customData);
			// post event to segment.io
			var _logger = new UsageLogger();
			_logger.PostException(ex);
#endif

			// show friendly message 
			if (showMessage)
				System.Windows.Forms.MessageBox.Show(
					"Unfortunately, an error occured in the Warrior add-in!\n" +
					"We have logged the error and will try to adress it as soon as possible.\n" +
					"In the mean time, if you have any questions or comments, please e-mail eric.selin@gmail.com");
		}
	}
}
