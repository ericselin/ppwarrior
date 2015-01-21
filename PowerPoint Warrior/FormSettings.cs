using System;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PowerPoint_Warrior
{
    public partial class FormSettings : Form
    {
        public FormSettings()
        {
            InitializeComponent();
            // Logging checkbox, email, and license
            chkLogging.Checked = Properties.Settings.Default.EnableLogging;
            tbEmail.Text = Properties.Settings.Default.UserEmail;
            tbLicenseKey.Text = Properties.Settings.Default.LicenseKey;
            // Edition text, if company name exists, append that as well
            lblVersion.Text = Properties.Settings.Default.Edition;
            if (!string.IsNullOrEmpty(Properties.Settings.Default.Company))
                lblVersion.Text = lblVersion.Text + " / " + Properties.Settings.Default.Company;
            // Valid until
            lblValidUntil.Text = Properties.Settings.Default.ValidUntil.ToString("d");
            // Give focus to e-mail
            tbEmail.Focus();
        }

        private async void btnOk_Click(object sender, EventArgs e)
        {
            if (ValidateChildren())
            {
                // If e-mail or license key changed, check online for new license
                if (!string.IsNullOrEmpty(tbLicenseKey.Text) &&
                    (Properties.Settings.Default.UserEmail != tbEmail.Text ||
                    Properties.Settings.Default.LicenseKey != tbLicenseKey.Text))
                {
                    // Check license online
                    // let user know we are doing something
                    var btnText = btnOk.Text;
                    btnOk.Text = "Checking...";
                    btnOk.Enabled = false;
                    Cursor = Cursors.WaitCursor;
                    // check awaiting
                    WarriorCommon.License license = await checkLicense();
                    // return to regular state
                    btnOk.Text = btnText;
                    btnOk.Enabled = true;
                    Cursor = Cursors.Default;

                    if (license != null)
                    {
                        Properties.Settings.Default.UserEmail = tbEmail.Text;
                        Properties.Settings.Default.LicenseKey = tbLicenseKey.Text;
                        Properties.Settings.Default.Edition = license.Edition;
                        Properties.Settings.Default.Company = license.Company;
                        Properties.Settings.Default.ValidUntil = license.ValidUntil;
                    }
                }
                // If license key removed, just warn the user but do nothing
                else if (string.IsNullOrEmpty(tbLicenseKey.Text) && !string.IsNullOrEmpty(Properties.Settings.Default.LicenseKey))
                {
                    // Show warning, but don't update settings
                    MessageBox.Show("It is not possible to remove the license key.\n" +
                        "License information has not been updated.");
                }
                // If we get here, either nothing changed or the user is in the trial
                else
                {
                    // In case the user is in trial, we update the email
                    Properties.Settings.Default.UserEmail = tbEmail.Text;
                }

                // These settings are not part of licensing (need to get online), so can be updated every time
                Properties.Settings.Default.EnableLogging = chkLogging.Checked;

                // Save settings
                Properties.Settings.Default.Save();

                this.Close();
            }
        }

        private async System.Threading.Tasks.Task<WarriorCommon.License> checkLicense()
        {
            var license = await WarriorCommon.LicenseManager.CheckLicense(tbEmail.Text, tbLicenseKey.Text);

            // Update settings accordingly
            if (license != null)
            {
                return license;
            }
            else
            {
                var result = MessageBox.Show("Could not verify license. Make sure your information is correct " +
                    "and that you are connected to the Internet.\n" +
                    "For assistance, please e-mail eric.selin@gmail.com",
                    "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                if (result == DialogResult.Retry)
                {
                    await checkLicense();
                }
                return null;
            }
        }

        private void tbEmail_Validating(object sender, CancelEventArgs e)
        {
            string errorMsg;
            if (!validEmail(tbEmail.Text, out errorMsg))
            {
                // Cancel the event and select the text to be corrected by the user.
                e.Cancel = true;
                tbEmail.Select(0, tbEmail.Text.Length);

                // Set the ErrorProvider error with the text to display.  
                this.errorpSettings.SetError(tbEmail, errorMsg);
            }
        }

        private void tbEmail_Validated(object sender, EventArgs e)
        {
            errorpSettings.SetError(tbEmail, "");
        }

        private bool validEmail(string emailAddress, out string errorMessage)
        {
            if (string.IsNullOrEmpty(emailAddress))
            {
                errorMessage = "E-mail required!";
                return false;
            }
            // Check email regex, from http://www.regular-expressions.info/email.html
            else if (!Regex.IsMatch(emailAddress, @"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}", RegexOptions.IgnoreCase))
            {
                errorMessage = "Please provide a valid e-mail address.";
                return false;
            }

            errorMessage = "";
            return true;
        }
    }
}
