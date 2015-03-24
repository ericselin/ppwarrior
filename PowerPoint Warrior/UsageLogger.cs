using Segment;
using Segment.Model;
using System;

namespace PowerPoint_Warrior
{
    public class UsageLogger
    {
        // segment.io write key
        private const string WRITE_KEY = "bjgcaeb4bj";
        // User email
        private string userId;
        // Traits to send with every event 
        private Traits traits;

        public UsageLogger()
        {
            // initialize segment.io, if debugging, do not use async
#if DEBUG
            Analytics.Initialize(WRITE_KEY, new Config().SetAsync(false));
            System.Diagnostics.Debug.Print("Debugging, calls to segment.io are synchronous");
#else
            Analytics.Initialize(WRITE_KEY);
#endif
            // set userId - if we don't have the email, just send the username
            this.userId = string.IsNullOrEmpty(Properties.Settings.Default.UserEmail) ? Information.GetUserName() : this.userId;
            // identify user
            UpdateIdentity();
        }

        internal void PostException(Exception ex)
        {
            var properties = new Segment.Model.Properties();
            properties.Add("Exception message", ex.Message);
            properties.Add("Stack trace", ex.StackTrace);
            // post to segment.io
            var options = new Options().SetContext(new Context().Add("traits", traits));
            Analytics.Client.Track(userId, "Encountered exception", properties, options);
        }

        public void PostUsage(string eventName, string buttonName)
        {
            // properties to include in track call
            var properties = new Segment.Model.Properties();
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
            UpdateIdentity();
        }

		public void UpdateIdentity(string userEmail, string company)
		{
			UpdateIdentity();
		}

		public void UpdateIdentity(string userEmail, string company, string something)
		{
			UpdateIdentity();
		}

		public void UpdateIdentity()
		{
			var userEmail = Properties.Settings.Default.UserEmail;
			var company = Properties.Settings.Default.Company;	
			// set traits
			traits = new Traits();
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
			// just in case, add windows user name
            traits.Add("Windows User Name", Information.GetUserName());
			// license
			traits.Add("Plan", Properties.Settings.Default.Edition);
			traits.Add("License Valid Until", Properties.Settings.Default.ValidUntil);
			// versions
            traits.Add("ClickOnce Version", Information.GetClickOnceVersion());
            traits.Add("Office Version", Globals.ThisAddIn.Application.Version);
            // Windows versions: http://msdn.microsoft.com/en-us/library/windows/desktop/ms724832%28v=vs.85%29.aspx
            traits.Add("Windows Version", Information.GetWindowsVersion());
            // if there is an email, identify
            if (!string.IsNullOrEmpty(userEmail))
            {
                Analytics.Client.Identify(this.userId, traits);
            }
            // if we have a company, also create a group - not used for intercom!
            /* TURN ON IF NEEDED
			if (!string.IsNullOrEmpty(company))
            {
                Analytics.Client.Group(this.userId, company, new Traits { { "name", company } }, 
					new Options().SetIntegration("Intercom", false) );
            } */
        }
    }
}
