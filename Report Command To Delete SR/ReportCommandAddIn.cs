using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using System;

namespace Report_Command_To_Delete_SR
{
    [AddIn("Report Command AddIn", Version = "1.0.0.0")]
    public class ReportCommandAddIn : IReportCommand2
    {
        #region IReportCommand Members
        static public IGlobalContext _globalContext;
        /// <summary>
        /// 
        /// </summary>
        public bool Enabled(IList<IReportRow> rows)
        {
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Execute(IList<IReportRow> rows)
        {
            List<int> incId = new List<int>();
            List<int> incVinId = new List<int>();
            int parentIncidentId = 0;// This will store Internal Incident ID when user delete Incident_VIN record

            foreach (IReportRow row in rows)
            {
                IList<IReportCell> cells = row.Cells;

                foreach (IReportCell cell in cells)
                {
                    // This condition is for parent report
                    if (cell.Name == "Internal_Incident_ID")
                    {
                        incId.Add(Convert.ToInt32(cell.Value));                        
                        break;
                    }
                    /* Below 2 check is for sub report */
                    if (cell.Name == "Incident_VIN_ID")
                    {
                        incVinId.Add(Convert.ToInt32(cell.Value));                        
                    }
                    if(cell.Name == "Parent_Incident")//meaning.. Incident ID mapped to Incident_VIN record
                    {
                        parentIncidentId = Convert.ToInt32(cell.Value);
                    }
                }
            }
            //If Incident ID exist 
            if(incId != null && incId.Count >0)
            {
                DialogResult result = MessageBox.Show("Are you sure you want to delete selected SR(s)?", "Confirmation", MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.Yes)
                {
                    RightNowConnectService.GetService().DeleteInternalIncident(incId);
                    _globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Save);
                }
            }
            //If incident_Vin id exist
            if(incVinId != null && incVinId.Count > 0)
            {
                DialogResult result = MessageBox.Show("Are you sure you want to delete selected VIN(s)?", "Confirmation", MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.Yes)
                {
                    RightNowConnectService.GetService().DeleteIncidentVIN(incVinId);
                    //If user has deleted all VIN of affected SR, then delete Internal incident too
                    RightNowConnectService.GetService().IncidentVinCountForInternalInc(parentIncidentId);
                    _globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Save);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public IList<RightNow.AddIns.Common.ReportRecordIdType> RecordTypes
        {
            get
            {
                IList<ReportRecordIdType> typeList = new List<ReportRecordIdType>();

                typeList.Add(ReportRecordIdType.Answer);
                typeList.Add(ReportRecordIdType.Chat);
                typeList.Add(ReportRecordIdType.CloudAcct2Search);
                typeList.Add(ReportRecordIdType.Contact);
                typeList.Add(ReportRecordIdType.ContactList);
                typeList.Add(ReportRecordIdType.Document);
                typeList.Add(ReportRecordIdType.Flow);
                typeList.Add(ReportRecordIdType.Incident);
                typeList.Add(ReportRecordIdType.Mailing);
                typeList.Add(ReportRecordIdType.MetaAnswer);
                typeList.Add(ReportRecordIdType.Opportunity);
                typeList.Add(ReportRecordIdType.Organization);
                typeList.Add(ReportRecordIdType.Question);
                typeList.Add(ReportRecordIdType.QueuedReport);
                typeList.Add(ReportRecordIdType.Quote);
                typeList.Add(ReportRecordIdType.QuoteProduct);
                typeList.Add(ReportRecordIdType.Report);
                typeList.Add(ReportRecordIdType.Segment);
                typeList.Add(ReportRecordIdType.Survey);
                typeList.Add(ReportRecordIdType.Task);
                typeList.Add(ReportRecordIdType.CustomObjectAll);
                return typeList;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Text
        {
            get
            {
                return "Delete";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "Delete/Remove Sales Release";
            }
        }

        public IList<string> CustomObjectRecordTypes
        {
            get
            {
                IList<string> typeList = new List<string>();

                typeList.Add("Incident_VIN");
                typeList.Add("Bus");
                return typeList;
            }
        }

        Image IReportCommand2.Image16
        {
            get
            {
                return Properties.Resources.AddIn16;
            }
        }

        Image IReportCommand2.Image32
        {
            get
            {
                return Properties.Resources.AddIn16;
            }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            _globalContext = GlobalContext;
            return true;
        }

        #endregion
    }
}
