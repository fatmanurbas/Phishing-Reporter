/*
  * licensed under the GNU General Public License v3.0
 */

using Microsoft.Office.Core;
using PhishingReporter.Properties;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Security.Cryptography;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Text;
using System.Net.Http;
using System.Threading.Tasks;


namespace PhishingReporter
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private Outlook.Application outlookApplication;
        public Ribbon(Outlook.Application application)
        {
            this.outlookApplication = application;
        }

        public Bitmap getGroup1Image(IRibbonControl control)
        {
            return Resources.phishing;
        }

        // Functions
        public void reportPhishing(Office.IRibbonControl control)
        {
            string userNote = "";
            string value = "Add a note";
            if (Tmp.InputBox("Report Mail", "Add some informations about this feedback. This field is optional", ref value) == DialogResult.OK)
            {
                if (value != "Add a note")
                    userNote = value;

                try
                {
                    // Fonksiyon çağrısı burada yapılır
                    reportPhishingEmailToSecurityTeamAsync(control, userNote);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("An error occured with the function: " + ex.Message);
                }

            }
        }

        private void reportPhishingEmailToSecurityTeamAsync(IRibbonControl control, string note)
        {
            Dictionary<string, object> senderMails = new Dictionary<string, object>();
            Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
            string reportedItemType = "NaN"; // email, contact, appointment ...etc

            if (selection.Count < 1) // no item is selected
            {
                MessageBox.Show("Select an email before reporting.", "Error");
            }
            else if (selection.Count > 1) // many items selected
            {
                MessageBox.Show("You can report 1 email at a time.", "Error");
            }
            else // only 1 item is selected
            {
                if (selection[1] is Outlook.MeetingItem || selection[1] is Outlook.ContactItem || selection[1] is Outlook.AppointmentItem || selection[1] is Outlook.TaskItem || selection[1] is Outlook.MailItem)
                {
                    // Identify the reported item type
                    if (selection[1] is Outlook.MeetingItem)
                        reportedItemType = "MeetingItem";
                    else if (selection[1] is Outlook.ContactItem)
                        reportedItemType = "ContactItem";
                    else if (selection[1] is Outlook.AppointmentItem)
                        reportedItemType = "AppointmentItem";
                    else if (selection[1] is Outlook.TaskItem)
                        reportedItemType = "TaskItem";
                    else if (selection[1] is Outlook.MailItem)
                        reportedItemType = "MailItem";

                    // Prepare Reported Email
                    Object mailItemObj = (selection[1] as object) as Object;
                    MailItem mailItem = (reportedItemType == "MailItem") ? selection[1] as MailItem : null; // If the selected item is an email

                    MailItem reportEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);

                    senderMails.Add("fromMail", mailItem.SenderEmailAddress == null ? "Draft" : mailItem.SenderEmailAddress);
                    senderMails.Add("userMail", GetCurrentUserInfos());
                    senderMails.Add("userNote", note);
                    senderMails.Add("htmlBody", mailItem.HTMLBody);

                    List<Dictionary<string, string>> attachments = new List<Dictionary<string, string>>();
                    Dictionary<string, string> attachment = new Dictionary<string, string>();

                    //TODO: Fatma, eğer mail eki mevcutsa, her bir ek için attachments objesine attachment eklenecek (for ile tüm ekler dönülmeli)
                    
                    /*
                    attachment["filename"] = "";
                    attachment["content"] = "";//base64 dosya içeriği
                    attachments.Add(attachment);
                    senderMails.Add("attachments", attachments);
                    */

                    using (var client1 = new HttpClient())
                    {
                        var endpoint = new Uri("https://m365.phishup.co/mail-provider-auth-service/microsoft365/user-mail-feedback");

                        string json = JsonConvert.SerializeObject(senderMails);
                        var payload = new StringContent(json, Encoding.UTF8, "application/json");

                        try
                        {
                            var response = client1.PostAsync(endpoint, payload).Result;
                            response.EnsureSuccessStatusCode();
                            string result = response.Content.ReadAsStringAsync().Result;

                            MessageBox.Show("Thank you for reporting.", "Thank you");
                        }
                        catch (HttpRequestException ex)
                        {
                            MessageBox.Show($"HTTP request failed: {ex.Message}", "Error");
                        }
                        catch (TaskCanceledException ex)
                        {

                            MessageBox.Show($"Task was canceled: {ex.Message}", "Error");
                        }
                        catch (System.Exception ex)
                        {

                            MessageBox.Show($"An error occurred: {ex.Message}", "Error");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("You cannot report this item", "Error");
                }
            }
        }

        public String GetCurrentUserInfos()
        {
            string str = "";

            Outlook.AddressEntry addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    str += currentUser.PrimarySmtpAddress;
                }
            }
            return str;
        }
        private void MoveMailToSpamFolder(MailItem mailItem)
        {
            Outlook.Application outlookApplication = Globals.ThisAddIn.Application;
            Outlook.Selection selectedItems = outlookApplication.ActiveExplorer().Selection;

            if (selectedItems != null && selectedItems.Count > 0)
            {
                Outlook.MAPIFolder spamFolder = outlookApplication.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);

                foreach (object selectedItem in selectedItems)
                {
                    if (selectedItem is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem2 = selectedItem as Outlook.MailItem;
                        mailItem2.Move(spamFolder);
                    }
                }
            }
        }


        #region IRibbonExtensibility Members
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PhishingReporter.Ribbon.xml");
        }
        #endregion

        #region Ribbon Callbacks
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        #endregion

        #region Helpers
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion
    }



    public static class MailItemExtensions
    {
        private const string HeaderRegex =
            @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)" +
                "(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string[] Headers(this MailItem mailItem, string name)
        {
            var headers = mailItem.HeaderLookup();
            if (headers.Contains(name))
                return headers[name].ToArray();
            return new string[0];
        }

        public static ILookup<string, string> HeaderLookup(this MailItem mailItem)
        {
            var headerString = mailItem.HeaderString();
            var headerMatches = Regex.Matches
                (headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();
            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value);
        }

        public static string HeaderString(this MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor
                .GetProperty(TransportMessageHeadersSchema);
        }

    }
}
