using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools
{
    [ComVisible(true)]
    public class OutlookToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private String strEmailAddress = "me@me.com";

        public OutlookToolsRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookTools.OutlookToolsRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public System.Drawing.Bitmap GetImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "textButtonReportSpam":
                    {
                        //add a icon for the button here
                        //return new System.Drawing.Bitmap(Properties.Resources.imgLogo);
                        return null;
                    }
            }
            return null;
        }

        #endregion

        #region "Button: Report Spam"

        public void OnReportSpamButton(Office.IRibbonControl control)
        {

            Outlook.Application application = new Outlook.Application();
            Outlook.NameSpace ns = application.GetNamespace("MAPI");

            //get selected outlook object / mail item
            Object selectedObject = application.ActiveExplorer().Selection[1];
            Outlook.MailItem selectedMail = (Outlook.MailItem)selectedObject;

            //compose a new message
            Outlook.MailItem newMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            newMail.Recipients.Add(strEmailAddress);
            newMail.Subject = "Spam email";

            //attach the selected mail item (spam to be sent)
            newMail.Attachments.Add(selectedMail, Microsoft.Office.Interop.Outlook.OlAttachmentType.olEmbeddeditem);

            //send the email
            newMail.Send();

            //delete the selected mail item (spam to be deleted)
            selectedMail.Delete();

        }

        #endregion

        #region "Get SMTP account methods "

        public Outlook.Account GetAccountForEmailAddress(Outlook.Application application,
            string smtpAddress)
        {
            // Loop over the Accounts collection of the current Outlook session.
            Outlook.Accounts accounts = application.Session.Accounts;

            foreach (Outlook.Account account in accounts)
            {
                // When the email address matches, return the account.
                if (account.SmtpAddress.ToLower() == smtpAddress.ToLower())
                {
                    return account;
                }
            }
            // If you get here, no matching account was found.
            throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!",
                smtpAddress));
        }


        public Outlook.Account GetDefaultAccount(Outlook.Application application)
        {

            // Get the Store for CurrentFolder.
            Outlook.Folder folder =
                application.ActiveExplorer().CurrentFolder
                as Outlook.Folder;
            Outlook.Store store = folder.Store;
            Outlook.Accounts accounts =
                application.Session.Accounts;
            // Enumerate accounts to find
            // account.DeliveryStore for store.
            foreach (Outlook.Account account in accounts)
            {
                if (account.DeliveryStore.StoreID ==
                    store.StoreID)
                {
                    return account;
                }
            }

            // If you get here, no matching account was found.
            throw new System.Exception(string.Format("No Account found!"));
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
}
