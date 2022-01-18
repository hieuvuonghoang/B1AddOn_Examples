
using FTIAddOn.SystemForms.SaleOrder;
using System;
using System.Diagnostics;

namespace FTIAddOn
{

    public class B1Events
    {
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        private const string SOURCE = "FTI AddOn";
        private const string LOG = "AddOn";

        public B1Events()
        {
            SetApplication();
            SetFilters();
            EventHandlers();

            if (!EventLog.SourceExists(SOURCE))
                EventLog.CreateEventSource(SOURCE, LOG);
            using (EventLog eventLog = new EventLog(LOG))
            {
                eventLog.Source = SOURCE;
                eventLog.WriteEntry("Application has started", EventLogEntryType.Information);
            }
        }

        private void SetApplication()
        {

            // *******************************************************************
            // Use an SboGuiApi object to establish connection
            // with the SAP Business One application and return an
            // initialized appliction object
            // *******************************************************************

            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();

            // by following the steped specified above the following
            // statment should be suficient for either development or run mode

            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

            // connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString);

            // get an initialized application object

            SBO_Application = SboGuiApi.GetApplication(-1);

            oCompany = SBO_Application.Company.GetDICompany();

            SBO_Application.SetStatusBarMessage("Connected!", SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        private void SetFilters()
        {
            SAPbouiCOM.EventFilters oFilters;
            SAPbouiCOM.EventFilter oFilter;

            // Create a new EventFilters object
            oFilters = new SAPbouiCOM.EventFilters();

            // add an event type to the container
            // this method returns an EventFilter object
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);

            // assign the form type on which the event would be processed
            oFilter.Add(139); // Orders Form

            SBO_Application.SetFilter(oFilters);
        }

        public void EventHandlers()
        {
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            //SBO_Application.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(SBO_Application_ProgressBarEvent);
            //SBO_Application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
            SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
            //SBO_Application.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(SetPrinter);
        }

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes eventType)
        {
            switch (eventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    if (!EventLog.SourceExists(SOURCE))
                        EventLog.CreateEventSource(SOURCE, LOG);
                    using (EventLog eventLog = new EventLog(LOG))
                    {
                        eventLog.Source = SOURCE;
                        eventLog.WriteEntry("FTI AddOn terminate", EventLogEntryType.Information);
                    }
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }

        private void SetPrinter(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //1282: Add
            //1283: Remove
            //1281: Find
            //1291: Last Record
            //1288: Next
            //1289: Previous
            //1290: First Record
            //1304: Refresh
            var oForm = SBO_Application.Forms.ActiveForm;
            switch (oForm.TypeEx)
            {
                case "139":
                    var congThucGiatCap = new CongThucGiatCap(oForm.UniqueID, SBO_Application, oCompany);
                    congThucGiatCap.MenuEvent(ref pVal, out BubbleEvent);
                    congThucGiatCap = null;
                    break;
            }
        }

        private static void ReportDataBefore(ref SAPbouiCOM.PrintEventInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        private void SBO_Application_ItemEvent(string formUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            switch(pVal.FormTypeEx)
            {
                case "139":
                    var congThucGiatCap = new CongThucGiatCap(formUID, SBO_Application, oCompany);
                    congThucGiatCap.ItemEvent(ref pVal, out BubbleEvent);
                    congThucGiatCap = null;
                    break;
            }
        }

        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            switch (BusinessObjectInfo.FormTypeEx)
            {
                case "139":
                    var congThucGiatCap = new CongThucGiatCap(BusinessObjectInfo.FormUID, SBO_Application, oCompany);
                    congThucGiatCap.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                    congThucGiatCap = null;
                    break;
            }
        }

        private void SBO_Application_ProgressBarEvent(ref SAPbouiCOM.ProgressBarEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        private void SBO_Application_StatusBarEvent(string Text, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {
        }

        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var oForm = SBO_Application.Forms.ActiveForm;
            switch (oForm.TypeEx)
            {
                case "139":
                    var congThucGiatCap = new CongThucGiatCap(eventInfo.FormUID, SBO_Application, oCompany);
                    congThucGiatCap.RightClick(ref eventInfo, out BubbleEvent);
                    congThucGiatCap = null;
                    break;
            }
        }

    }

}
