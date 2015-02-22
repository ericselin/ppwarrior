using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace PowerPoint_Warrior
{
    public static class Email
    {
        public static void SendFeedback(string subject)
        {
            Outlook.Application outlook = null;

            // Get Outlook application object
            try
            {
                object app = Marshal.GetActiveObject("Outlook.Application");
                outlook = app as Outlook.Application;
            }
            catch (COMException)
            {
                // This means that Outloook wasn't running, just let exception slide
            }

            // If outlook instance not found, create new instance
            if (outlook == null)
                outlook = new Outlook.Application();

            // If no oulook object, show exception and return
            if (outlook == null)
            {
                System.Windows.Forms.MessageBox.Show("Could not find Outlook or create new Outlook instance.\nPlease write feedback to eric.selin@gmail.com.");
                return;
            }

			Outlook.MailItem eMail =
                (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = subject;
            eMail.To = "eric@ppwarrior.com";
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            eMail.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
            eMail.Body = "Please write feedback here.\nBug reports, new feature requests and all sorts of other feedback welcome!";

			eMail.Display();
        }
    }
}

