using Segment;
using Segment.Model;
using System;

namespace WarriorCommon
{
    public class UsageLogger
    {
        // segment.io write key
        private const string WRITE_KEY = "bjgcaeb4bj";
        // User email
        private string userEmail;
        private string officeVersion;
        private string windowsVersion;

        public UsageLogger(string officeVersion, string userEmail)
        {
            // initialize segment.io, if debugging, do not use async
#if DEBUG
            Analytics.Initialize(WRITE_KEY, new Config().SetAsync(false));
            System.Diagnostics.Debug.Print("Debugging, calls to segment.io are synchronous");
#else
            Analytics.Initialize(WRITE_KEY);
#endif
            // set user email and versions
            this.userEmail = userEmail;
            this.officeVersion = officeVersion;
            this.windowsVersion = string.Format("{0}.{1}", Environment.OSVersion.Version.Major, Environment.OSVersion.Version.Minor);
            // identify user
            UpdateIdentity(userEmail);
        }

        public void PostUsage(string eventName, string buttonName)
        {
            // properties to include in track call
            var properties = new Properties();
            // if there is a button name, include it in properties
            if (buttonName != null && buttonName.Length > 0)
                properties.Add("Button Name", buttonName);
            // if we don't have the email, just send the username
            var userId = string.IsNullOrEmpty(userEmail) ? Information.GetUserName() : userEmail;
            // post to segment.io
            Analytics.Client.Track(userId, eventName, properties);
        }

        public void PostUsage(string eventName)
        {
            PostUsage(eventName, null);
        }

        public void UpdateIdentity(string userEmail)
        {
            // if there is an email, identify
            if (!string.IsNullOrEmpty(userEmail))
            {
                this.userEmail = userEmail;
                Analytics.Client.Identify(this.userEmail, new Traits()
                {
                    { "Assembly Version", Information.GetAssemblyVersion() },
                    { "ClickOnce Version", Information.GetClickOnceVersion() },
                    { "Office Version", officeVersion },
                    // Windows versions: http://msdn.microsoft.com/en-us/library/windows/desktop/ms724832%28v=vs.85%29.aspx
                    { "Windows Version", windowsVersion },
                    { "email", this.userEmail }
                }); 
            }
        }
    }
}
