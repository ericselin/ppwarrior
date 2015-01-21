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
        private string userId;
        // Version info (need to get these from the addin)
        private string officeVersion;
        private string windowsVersion;
        // Traits to send with every event 
        private Traits traits;

        public UsageLogger(string officeVersion, string userEmail, string company)
        {
            // initialize segment.io, if debugging, do not use async
#if DEBUG
            Analytics.Initialize(WRITE_KEY, new Config().SetAsync(false));
            System.Diagnostics.Debug.Print("Debugging, calls to segment.io are synchronous");
#else
            Analytics.Initialize(WRITE_KEY);
#endif
            // set email - if we don't have the email, just send the username
            this.userId = string.IsNullOrEmpty(userEmail) ? Information.GetUserName() : this.userId;
            // set version info so we don't need to get it every time updateidentity is called
            this.officeVersion = officeVersion;
            this.windowsVersion = string.Format("{0}.{1}", Environment.OSVersion.Version.Major, Environment.OSVersion.Version.Minor);
            // identify user
            UpdateIdentity(userEmail, company);
        }

        internal void PostException(Exception ex)
        {
            var properties = new Properties();
            properties.Add("Exception message", ex.Message);
            properties.Add("Stack trace", ex.StackTrace);
            // post to segment.io
            var options = new Options().SetContext(new Context().Add("traits", traits));
            Analytics.Client.Track(userId, "Encountered exception", properties, options);
        }

        public void PostUsage(string eventName, string buttonName)
        {
            // properties to include in track call
            var properties = new Properties();
            // if there is a button name, include it in properties
            if (buttonName != null && buttonName.Length > 0)
                properties.Add("Button Name", buttonName);
            // post to segment.io
            var options = new Options().SetContext(new Context().Add("traits", traits));
            Analytics.Client.Track(userId, eventName, properties, options);
        }

        public void PostUsage(string eventName)
        {
            PostUsage(eventName, null);
        }

        public void UpdateIdentity(string userEmail)
        {
            UpdateIdentity(userEmail, null);
        }

        public void UpdateIdentity(string userEmail, string company)
        {
            // set traits
            traits = new Traits();
            traits.Add("Assembly Version", Information.GetAssemblyVersion());
            traits.Add("ClickOnce Version", Information.GetClickOnceVersion());
            traits.Add("Office Version", officeVersion);
            // Windows versions: http://msdn.microsoft.com/en-us/library/windows/desktop/ms724832%28v=vs.85%29.aspx
            traits.Add("Windows Version", windowsVersion);
            // add email as trait
            if (!string.IsNullOrEmpty(userEmail))
            {
                this.userId = userEmail;
                traits.Add("Email", this.userId);
            }
            // add company as trait
            if (!string.IsNullOrEmpty(company))
            {
                traits.Add("Company", company);
            }
            // if there is an email, identify
            if (!string.IsNullOrEmpty(userEmail))
            {
                Analytics.Client.Identify(this.userId, traits);
            }
            // if we have a company, also create a group
            if (!string.IsNullOrEmpty(company))
            {
                Analytics.Client.Group(this.userId, company, new Traits { { "name", company } } );
            }
        }
    }
}
