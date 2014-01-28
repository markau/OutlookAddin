using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookTools
{
    public partial class ThisAddIn
    {

        //Outlook.Inspectors inspectors;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            /*
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            */
         }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /*
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";

                    //Outlook.MailItem mailItem2 = new Application.CreateItem();

                }

            }
        }
        */

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutlookTools.OutlookToolsRibbon();
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
