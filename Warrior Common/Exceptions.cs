using Mindscape.Raygun4Net;
using Mindscape.Raygun4Net.Messages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WarriorCommon
{
	public static class Exceptions
	{
		public static void Handle(Exception ex, string officeVersion, string userEmail, bool showMessage = true)
		{
#if !DEBUG
            // create raygun client and set appropriate flags
            RaygunClient _client = new RaygunClient("z2DRwwFDwxCn/6WrH/Irqg==");
            _client.User = Information.GetUserName();
            if (!string.IsNullOrEmpty(userEmail))
            {
                RaygunIdentifierMessage userInfo = new RaygunIdentifierMessage(_client.User);
                userInfo.Email = userEmail;
                _client.UserInfo = userInfo;
            }
            _client.ApplicationVersion = Information.GetAssemblyVersion();
            var customData = new Dictionary<string, string>() {
            { "Office Version", officeVersion }, { "ClickOnce Version", Information.GetClickOnceVersion() }, { "Assembly Version", Information.GetAssemblyVersion() } };
            // send exception to raygun
            _client.Send(ex, null, customData);
            // post event to segment.io
            var _logger = new UsageLogger(officeVersion, userEmail, null);
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
