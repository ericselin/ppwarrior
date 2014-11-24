using System;
using System.ComponentModel;
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
            chkLogging.Checked = Properties.Settings.Default.EnableLogging;
            tbEmail.Text = Properties.Settings.Default.UserEmail;
            tbEmail.Focus();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (ValidateChildren())
            {
                Properties.Settings.Default.EnableLogging = chkLogging.Checked;
                Properties.Settings.Default.UserEmail = tbEmail.Text;
                Properties.Settings.Default.Save();
                this.Close(); 
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
