using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn_TEST
{
    public partial class ThisAddIn
    {
        private Outlook.Items calendarItems;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Outlook.MAPIFolder calendarFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            calendarItems = calendarFolder.Items;
            calendarItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(calendarItems_ItemAdd);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }
        private void calendarItems_ItemAdd(object Item)
        {
            if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = Item as Outlook.AppointmentItem;
                if (appointment.MessageClass == "IPM.Schedule.Meeting.Request")
                {
                    // Check if the appointment was sent from the desired person
                    Outlook.Recipient desiredRecipient = appointment.Recipients["Name"];
                    if (desiredRecipient != null && desiredRecipient.Resolve())
                    {
                        // Delete the appointment from the calendar
                        appointment.Delete();
                    }
                }
            }
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
