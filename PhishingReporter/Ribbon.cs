/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
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
using HtmlAgilityPack;
using System.Collections.Generic;
using Newtonsoft.Json;
using Microsoft.Office.Tools;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PhishingReporter
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private ThisAddIn addIn;
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
                    reportPhishingEmailToSecurityTeam(control, userNote);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("An error occured with the function: " + ex.Message);
                }

            }
        }

        /*
         *  Helper functions 
         */

        private void reportPhishingEmailToSecurityTeam(IRibbonControl control, string note)
        {
            Dictionary<string, object> senderMails = new Dictionary<string, object>();
            Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
            string reportedItemType = "NaN"; // email, contact, appointment ...etc
            string reportedItemHeaders = "NaN";

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
                    {
                        reportedItemType = "MeetingItem";
                    }
                    else if (selection[1] is Outlook.ContactItem)
                    {
                        reportedItemType = "ContactItem";
                    }
                    else if (selection[1] is Outlook.AppointmentItem)
                    {
                        reportedItemType = "AppointmentItem";
                    }
                    else if (selection[1] is Outlook.TaskItem)
                    {
                        reportedItemType = "TaskItem";
                    }
                    else if (selection[1] is Outlook.MailItem)
                    {
                        reportedItemType = "MailItem";
                    }

                    // Prepare Reported Email
                    Object mailItemObj = (selection[1] as object) as Object;
                    MailItem mailItem = (reportedItemType == "MailItem") ? selection[1] as MailItem : null; // If the selected item is an email

                    MailItem reportEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                    reportEmail.Attachments.Add(selection[1] as Object);

                    try
                    {

                        reportEmail.To = Properties.Settings.Default.infosec_email;
                        reportEmail.Subject = (reportedItemType == "MailItem") ? "[POTENTIAL PHISH] " + mailItem.Subject : "[POTENTIAL PHISH] " + reportedItemType; // If reporting email, include subject; otherwise, state the type of the reported item

                        // Get Email Headers
                        if (reportedItemType == "MailItem")
                        {
                            reportedItemHeaders = mailItem.HeaderString();
                        }
                        else
                        {
                            reportedItemHeaders = "";
                        }
                         
                        senderMails.Add("fromMail", mailItem.SenderEmailAddress == null ? "Draft" : mailItem.SenderEmailAddress);
                        senderMails.Add("userMail", GetCurrentUserInfos());
                        senderMails.Add("userNote", note);
                        senderMails.Add("htmlBody", mailItem.HTMLBody);
                        if(reportedItemHeaders != "")
                        {
                            senderMails.Add("header", reportedItemHeaders);
                        }

                        List<Dictionary<string, string>> attachments = new List<Dictionary<string, string>>();
                        Dictionary<string, string> attachment = new Dictionary<string, string>();
                        if(mailItem.Attachments != null && mailItem.Attachments.Count > 0)
                        {

                        foreach (Attachment a in mailItem.Attachments)
                        {

                            var tempFilePath = Path.GetTempFileName();
                            a.SaveAsFile(tempFilePath);

                            // Read the content of the temporary file as bytes
                            byte[] attachmentBytes = File.ReadAllBytes(tempFilePath);

                            // Convert attachment content to Base64 string
                            string base64String = Convert.ToBase64String(attachmentBytes);

                            // Now you can use the base64String variable to send or process the attachment content
                            Console.WriteLine("Attachment Content (Base64): " + base64String);

                            // Delete the temporary file
                            File.Delete(tempFilePath);

                            attachment["filename"] = a.FileName;
                            attachment["content"] = base64String;//base64 dosya içeriği
                            attachments.Add(attachment);
                        }
                            senderMails.Add("attachments", attachments);
                        }
                       
                        // Prepare the email body

                        reportEmail.Body += JsonConvert.SerializeObject(senderMails);
                            reportEmail.Body += "\n";
                        MoveMailToSpamFolder(mailItem);

                        reportEmail.Save();
                        //reportEmail.Display(); // Helps in debugginng
                       reportEmail.Send(); // Automatically send the email
                       

                        // Enable if you want a second popup for confirmation
                        // MessageBox.Show("Thank you for reporting. We will review this report soon. - Information Security Team", "Thank you");


                        // Delete the reported email
                        //  mailItem.Delete();

                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("There was an error! An automatic email was sent to the support to resolve the issue.", "Do not worry");

                        MailItem errorEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                        errorEmail.To = Properties.Settings.Default.support_email;
                        errorEmail.Subject = "[Outlook Addin Error]";
                        errorEmail.Body = ("Addin error message: " + ex);
                        errorEmail.Save();
                        //errorEmail.Display(); // Helps in debugginng
                        errorEmail.Send(); // Automatically send the email
                    }
                }
                else
                {
                    MessageBox.Show("You cannot report this item", "Error");
                }
            }
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

        public String GetURLsAndAttachmentsInfo(MailItem mailItem)
        {
            string urls_and_attachments = "---------- URLs and Attachments ----------";

            var domainsInEmail = new List<string>();

            var emailHTML = mailItem.HTMLBody;
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(emailHTML);

            // extracting all links
            var urlsText = "";
            var urlNodes = doc.DocumentNode.SelectNodes("//a[@href]");
           

            // Get domains
            domainsInEmail = domainsInEmail.Distinct().ToList();
            urls_and_attachments += "\n # of unique Domains: " + domainsInEmail.Count;
            foreach (string item in domainsInEmail)
            {
                urls_and_attachments += "\n --> Domain: " + item.Replace(":", "[:]");
            }

            // Add Urls
            

            urls_and_attachments += "\n\n # of Attachments: " + mailItem.Attachments.Count;
            foreach (Attachment a in mailItem.Attachments)
            {
                
                var tempFilePath = Path.GetTempFileName();
                a.SaveAsFile(tempFilePath);

                // Read the content of the temporary file as bytes
                byte[] attachmentBytes = File.ReadAllBytes(tempFilePath);

                // Convert attachment content to Base64 string
                string base64String = Convert.ToBase64String(attachmentBytes);

                // Now you can use the base64String variable to send or process the attachment content
                Console.WriteLine("Attachment Content (Base64): " + base64String);

                // Delete the temporary file
                File.Delete(tempFilePath);
                urls_and_attachments += "\n --> Attachment: " + a.FileName + " "  + "\n" + "content base64:" + base64String;
            }
            return urls_and_attachments;
        }



        public String GetPluginDetails()
        {
            string pluginDetails = "---------- Report Phishing Plugin ----------";
            pluginDetails += "\n - Version: " + Properties.Settings.Default.plugin_version;
            pluginDetails += "\n - Usage: Report phishing emails to the Information Security Team.";
            pluginDetails += "\n - Support: " + Properties.Settings.Default.support_email;

            pluginDetails += "\n - Developer: Abdulla Albreiki (aalbraiki@hotmail.com)"; // You may delete this line if you like :)
            return pluginDetails;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PhishingReporter.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

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

        static string CalculateMD5(string filename)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(filename))
                {
                    var hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
        }
        private string GetHashSha256(string filename)
        {
            using (FileStream stream = File.OpenRead(filename))
            {
                SHA256Managed sha = new SHA256Managed();
                byte[] shaHash = sha.ComputeHash(stream);
                string result = "";
                foreach (byte b in shaHash) result += b.ToString("x2");
                return result;
            }
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
