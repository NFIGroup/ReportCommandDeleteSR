using System;
using System.Collections.Generic;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using Report_Command_To_Delete_SR.RightNowService;

namespace Report_Command_To_Delete_SR
{
    class RightNowConnectService
    {
        private static RightNowConnectService _rightnowConnectService;
        private static object _sync = new object();
        private static RightNowSyncPortClient _rightNowClient;
        public static IRecordContext _recordContext;
        public static IIncident _incidentRecord;
        public List<RNObject> _incidentVINObjects = new List<RNObject>();


        private RightNowConnectService()
        {

        }
        public static RightNowConnectService GetService()
        {
            if (_rightnowConnectService != null)
            {
                return _rightnowConnectService;
            }

            try
            {
                lock (_sync)
                {
                    if (_rightnowConnectService == null)
                    {
                        // Initialize client with current interface soap url 
                        string url = ReportCommandAddIn._globalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap);
                        EndpointAddress endpoint = new EndpointAddress(url);

                        BasicHttpBinding binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportWithMessageCredential);
                        binding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;

                        // Optional depending upon use cases
                        binding.MaxReceivedMessageSize = 1024 * 1024;
                        binding.MaxBufferSize = 1024 * 1024;
                        binding.MessageEncoding = WSMessageEncoding.Mtom;

                        _rightNowClient = new RightNowSyncPortClient(binding, endpoint);

                        BindingElementCollection elements = _rightNowClient.Endpoint.Binding.CreateBindingElements();
                        elements.Find<SecurityBindingElement>().IncludeTimestamp = false;
                        _rightNowClient.Endpoint.Binding = new CustomBinding(elements);
                        ReportCommandAddIn._globalContext.PrepareConnectSession(_rightNowClient.ChannelFactory);

                        _rightnowConnectService = new RightNowConnectService();
                    }

                }
            }
            catch (Exception e)
            {
                _rightnowConnectService = null;
                MessageBox.Show(e.Message);
            }
            return _rightnowConnectService;
        }

        /// <summary>
        /// Return individual fields as per query
        /// </summary>
        /// <param name="ApplicationID"></param>
        /// <param name="Query"></param>
        /// <returns> array of string delimited by '~'</returns>
        private string[] GetRNData(string ApplicationID, string Query)
        {
            string[] rnData = null;
            ClientInfoHeader hdr = new ClientInfoHeader() { AppID = ApplicationID };

            byte[] output = null;
            CSVTableSet data = null;

            try
            {
                data = _rightNowClient.QueryCSV(hdr, Query, 1000, "~", false, false, out output);
                string dataRow = String.Empty;
                if (data != null && data.CSVTables.Length > 0 && data.CSVTables[0].Rows.Length > 0)
                {
                    return data.CSVTables[0].Rows;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return rnData;
        }

        /// <summary>
        /// Check if Internal incident has any Incident_VIN child record, if not then delete internal incident
        /// </summary>
        /// <param name="internalIncID"> Reported Incident Org ID</param>
        /// <returns> </returns>
        public void IncidentVinCountForInternalInc(int internalIncID)
        {
            string query = "SELECT count(ID) as count FROM CO.Incident_VIN WHERE incident = " + internalIncID;
            string[] resultSet = GetRNData("Get incident_VIN count", query);
            if (resultSet != null && resultSet.Length > 0)
            {
                if(resultSet[0] == "0")//if count is 0 then delete internal incident too
                {
                    List<int> incId = new List<int>();
                    incId.Add(internalIncID);
                    DeleteInternalIncident(incId);
                }
            }
            return ;
        }
        /// <summary>
        /// Delete Internal Incident Records
        /// </summary>
        /// <param name="deleteIncIDs"></param>
        /// <returns></returns>
        public void DeleteInternalIncident(List<int> deleteIncIDs)
        {
            try
            {
                List<RNObject> deleteObject = new List<RNObject>();
                for (int i = 0; i < deleteIncIDs.Count; i++)
                {

                    Incident incObj = new Incident();
                    incObj.ID = new ID
                    {
                        id = deleteIncIDs[i],
                        idSpecified = true
                    };

                    deleteObject.Add(incObj);
                }

                //BatchResponseItem[] batchRes = rspc.Batch(clientInfoHeader, requestItems);
                if (callBatchJob(getDestroyMsg(deleteObject)))
                {
                    //Once SR/internal Incident is deleted, decrement the SR number in no_of_sr field of Reported Incident
                    string oldSRNo = getFieldFromIncidentRecord("CO", "no_of_sr");//old sr number
                    int currentSrNo = (oldSRNo != "") ? Convert.ToInt32(oldSRNo) - deleteObject.Count : 0;//- no of SR selected for deletion
                    setIncidentField("CO", "no_of_sr", currentSrNo.ToString());//update new SR num
                }
            }
            catch (Exception ex)
            {
                //SR mean internal incident
                MessageBox.Show("Exception in Deleting SR record: " + ex.Message);
            }
            return;
        }
        /// <summary>
        /// Delete Incident VIN Records
        /// </summary>
        /// <param name="incVinID"></param>
        /// <returns></returns>
        public void DeleteIncidentVIN(List<int> deleteVins)
        {
            try
            {
                List<RNObject> deleteObject = new List<RNObject>();
                for (int i = 0; i < deleteVins.Count; i++)
                {

                    GenericObject genObj = new GenericObject
                    {
                        ObjectType = new RNObjectType
                        {
                            Namespace = "CO",
                            TypeName = "Incident_VIN"
                        }
                    };
                    genObj.ID = new ID
                    {
                        id = deleteVins[i],
                        idSpecified = true
                    };

                    deleteObject.Add(genObj);
                }
                //BatchResponseItem[] batchRes = rspc.Batch(clientInfoHeader, requestItems);
                callBatchJob(getDestroyMsg(deleteObject));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception in Deleting Incident_VIN record: " + ex.Message);
            }
            return;
        }
        /// <summary>
        /// Method which is called to get value of a custom field of Incident record.
        /// </summary>
        /// <param name="packageName">The name of the package.</param>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public string getFieldFromIncidentRecord(string packageName, string fieldName)
        {
            string value = "";
            _recordContext = ReportCommandAddIn._globalContext.AutomationContext.CurrentWorkspace;
            _incidentRecord = (IIncident)_recordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
            IList<ICustomAttribute> incCustomAttributes = _incidentRecord.CustomAttributes;

            foreach (ICustomAttribute val in incCustomAttributes)
            {
                if (val.PackageName == packageName)//if package name matches
                {
                    if (val.GenericField.Name == packageName + "$" + fieldName)//if field matches
                    {
                        if (val.GenericField.DataValue.Value != null)
                        {
                            value = val.GenericField.DataValue.Value.ToString();
                            break;
                        }
                    }
                }
            }
            return value;
        }
        /// <summary>
        /// Method which is use to set incident field 
        /// </summary>
        /// <param name="pkgName">package name of custom field</param>
        /// <param name="fieldName">field name</param>
        /// <param name="value">value of field</param>
        public void setIncidentField(string pkgName, string fieldName, string value)
        {

            IList<ICustomAttribute> incCustomAttributes = _incidentRecord.CustomAttributes;

            foreach (ICustomAttribute val in incCustomAttributes)
            {
                if (val.PackageName == pkgName)
                {
                    if (val.GenericField.Name == pkgName + "$" + fieldName)
                    {
                        switch (val.GenericField.DataType)
                        {
                            case RightNow.AddIns.Common.DataTypeEnum.INTEGER:
                                if (value.Trim() == "" || value.Trim() == null)
                                {
                                    val.GenericField.DataValue.Value = null;
                                }
                                else
                                {
                                    val.GenericField.DataValue.Value = Convert.ToInt32(value);
                                }
                                break;
                        }
                    }
                }
            }
            return;
        }
        /// <summary>
        /// Create DestroyMsg object
        /// </summary>
        /// <param name="coList">RNObject List</param>
        /// <returns> DestroyMsg</returns>
        private DestroyMsg getDestroyMsg(List<RNObject> coList)
        {
            DestroyMsg deleteMsg = new DestroyMsg();
            DestroyProcessingOptions deleteProcessingOptions = new DestroyProcessingOptions();
            deleteProcessingOptions.SuppressExternalEvents = true;
            deleteProcessingOptions.SuppressRules = true;
            deleteMsg.ProcessingOptions = deleteProcessingOptions;

            deleteMsg.RNObjects = coList.ToArray();

            return deleteMsg;
        }
        /// <summary>
        /// Perform Batch operation
        /// </summary>
        /// <param name="msg">BatchRequestItem Item</param>
        public bool callBatchJob(Object msg)
        {
            try
            {
                /*** Form BatchRequestItem structure ********************/

                BatchRequestItem[] requestItems = new BatchRequestItem[1];

                BatchRequestItem requestItem = new BatchRequestItem();

                requestItem.Item = msg;

                requestItems[0] = requestItem;
                requestItems[0].CommitAfter = true;
                requestItems[0].CommitAfterSpecified = true;
                /*********************************************************/

                ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
                clientInfoHeader.AppID = "Batcher";

                BatchResponseItem[] batchRes = _rightNowClient.Batch(clientInfoHeader, requestItems);
                //If response type is RequestErrorFaultType then show the error msg 
                if (batchRes[0].Item.GetType().Name == "RequestErrorFaultType")
                {
                    RequestErrorFaultType requestErrorFault = (RequestErrorFaultType)batchRes[0].Item;
                    MessageBox.Show("There is an error with batch job :: " + requestErrorFault.exceptionMessage);
                    return false;
                }
            }
            catch (FaultException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }
    }
}
