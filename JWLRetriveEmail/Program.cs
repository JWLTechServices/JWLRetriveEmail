using EAGetMail;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace JWLRetriveEmail
{
    public static class Program
    {
        static void Main(string[] args)
        {
            //  DataTable dt = new DataTable();
            // ProcessBBBDELfileToDataTable("", dt);
            // return;

            StartProcessing();
            clsCommon objCommon = new clsCommon();
            string appName = objCommon.GetConfigValue("ApplicationName");
            if (clsCommon.IsException)
            {
                string emailSubject = "Got Exception while running " + appName + " on " + DateTime.Now.ToString("yyyyMMdd");
                string emailBody = emailSubject + System.Environment.NewLine + "Requesting you to please go and check error log file for :" + DateTime.Now.ToString("yyyyMMdd");
                objCommon.SendExceptionMail(emailSubject, emailBody);
            }
        }

        private static void StartProcessing()
        {
            clsCommon objCommon = new clsCommon();
            string appName = objCommon.GetConfigValue("ApplicationName");
            string emailSubject = "";
            try
            {
                string executionLogMessage;
                executionLogMessage = "Beginning the new instance for " + appName + " processing ";
                objCommon.WriteExecutionLog(executionLogMessage);

                MailServer oServer = new MailServer(objCommon.GetConfigValue("IMAPHost"), objCommon.GetConfigValue("IMAPUser"), objCommon.GetConfigValue("IMAPPassword"),
                             ServerProtocol.Imap4);

                // Enable SSL connection.
                oServer.SSLConnection = true;

                // Set 993 SSL port
                oServer.Port = Convert.ToInt32(objCommon.GetConfigValue("IMAPPort"));  //993;

                // MailClient oClient = new MailClient("TryIt");
                MailClient oClient = new MailClient(objCommon.GetConfigValue("EAGetMailLicenseCode"));
                oClient.Connect(oServer);

                // retrieve unread/new email only
                oClient.GetMailInfosParam.Reset();
                oClient.GetMailInfosParam.GetMailInfosOptions = GetMailInfosOptionType.NewOnly;


                MailInfo[] infos = oClient.GetMailInfos();

                executionLogMessage = "Number of unseen emails found : " + infos.Length;
                objCommon.WriteExecutionLog(executionLogMessage);


                // Console.WriteLine("Total {0} email(s)\r\n", infos.Length);
                for (int ia = 0; ia < infos.Length; ia++)
                {
                    MailInfo info = infos[ia];
                    // Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                    //     info.Index, info.Size, info.UIDL);

                    // Receive email from IMAP4 server
                    Mail oMail = oClient.GetMail(info);

                    string ReadEmailOnlyFrom = objCommon.GetConfigValue("ReadEmailOnlyFrom");
                    if (ReadEmailOnlyFrom.Split(',').Contains(oMail.From.Address.Split('@')[1]))
                    {
                        //string BBBReadEmailDomain = objCommon.GetConfigValue("BBBReadEmailDomain");
                        //string MXDRYDReadEmailDomain = objCommon.GetConfigValue("MXDRYDReadEmailDomain");
                        //string TGTReadEmailDomain = objCommon.GetConfigValue("TGTReadEmailDomain");
                        //string LBRANDSReadEmailDomain = objCommon.GetConfigValue("LBRANDSReadEmailDomain");
                        //string CDSReadEmailDomain = objCommon.GetConfigValue("CDSReadEmailDomain");
                        //string BBWReadEmailDomain = objCommon.GetConfigValue("BBWReadEmailDomain");

                        //if (BBBReadEmailDomain.Split(',').Contains(oMail.From.Address.Split('@')[1]) || MXDRYDReadEmailDomain.Split(',').Contains(oMail.From.Address.Split('@')[1]) || TGTReadEmailDomain.Split(',').Contains(oMail.From.Address.Split('@')[1])
                        //     || LBRANDSReadEmailDomain.Split(',').Contains(oMail.From.Address.Split('@')[1])
                        //     || CDSReadEmailDomain.Split(',').Contains(oMail.From.Address.Split('@')[1])
                        //     || BBWReadEmailDomain.Split(',').Contains(oMail.From.Address.Split('@')[1]))
                        // {
                        string BBBReadEmailSubject = objCommon.GetConfigValue("BBBReadEmailSubject");
                        string MXDRYDReadEmailSubject = objCommon.GetConfigValue("MXDRYDReadEmailSubject");
                        string TGTReadEmailSubject = objCommon.GetConfigValue("TGTReadEmailSubject");
                        string LBRANDSReadEmailSubject = objCommon.GetConfigValue("LBRANDSReadEmailSubject");
                        string CDSReadEmailSubject = objCommon.GetConfigValue("CDSReadEmailSubject");
                        string BBWReadEmailSubject = objCommon.GetConfigValue("BBWReadEmailSubject");
                        string ROVCONReadEmailSubject = objCommon.GetConfigValue("ROVCONReadEmailSubject");
                        string BLUDOTReadEmailSubject = objCommon.GetConfigValue("BLUDOTReadEmailSubject");

                        if ((oMail.Subject.Contains(BBBReadEmailSubject) || oMail.Subject.Contains(MXDRYDReadEmailSubject)
                            || oMail.Subject.Contains(TGTReadEmailSubject) || oMail.Subject.Contains(LBRANDSReadEmailSubject)
                            || oMail.Subject.Contains(CDSReadEmailSubject) || oMail.Subject.Contains(BBWReadEmailSubject)
                            || oMail.Subject.Contains(ROVCONReadEmailSubject) || oMail.Subject.Contains(ROVCONReadEmailSubject)))
                        {
                            //  continue;
                            executionLogMessage = "ReadEmailOnlyFrom :  " + oMail.From.Address.ToString() + System.Environment.NewLine;
                            objCommon.WriteExecutionLog(executionLogMessage);

                            executionLogMessage = "From :  " + oMail.From.ToString() + System.Environment.NewLine;
                            executionLogMessage += "Email Address :  " + oMail.From.Address.ToString() + System.Environment.NewLine;
                            executionLogMessage += "Subject : " + oMail.Subject + System.Environment.NewLine;
                            executionLogMessage += "ReceivedDate : " + oMail.ReceivedDate + System.Environment.NewLine;
                            objCommon.WriteExecutionLog(executionLogMessage);

                            Attachment[] atts = oMail.Attachments;
                            int count = atts.Length;

                            bool IsSeen = false;

                            for (int j = 0; j < count; j++)
                            {
                                objCommon.CleanAttachmentWorkingFolder();
                                Attachment att = atts[j];
                                string fileName = att.Name;

                                executionLogMessage = "Attachment Name : " + att.Name + System.Environment.NewLine;
                                objCommon.WriteExecutionLog(executionLogMessage);

                                //  string strExtension = att.Name;
                                try
                                {
                                    FileInfo fi = new FileInfo(fileName);
                                    string extn = fi.Extension;

                                    //.xls;
                                    if (extn.ToLower().Contains(".csv") || extn.ToLower().Contains(".xlsx") || extn.ToLower().Contains(".xls") || extn.ToLower().Contains(".bbb"))
                                    {
                                        string customerName = fileName.Split('-')[0].ToUpper();
                                        string locationCode = fileName.Split('-')[1].ToUpper();
                                        string productCode = fileName.Split('-')[2].ToUpper();
                                        string productSubCode = "";
                                        if (customerName == objCommon.GetConfigValue("BBBCustomerName") || customerName == objCommon.GetConfigValue("TGTCustomerName"))
                                        {
                                            productSubCode = fileName.Split('-')[3].ToUpper();
                                        }
                                        clsCommon.DSResponse objCMDsResponse = new clsCommon.DSResponse();
                                        objCMDsResponse = objCommon.GetReadEmailMappingDetails(customerName, locationCode, productCode);
                                        if (objCMDsResponse.dsResp.ResponseVal)
                                        {
                                            string tempfilePath = objCommon.GetConfigValue("AttachmentWorkingFolder");
                                            string attachmentPath = objCommon.GetConfigValue("AutomationFileLocation");
                                            string filePath = Convert.ToString(objCMDsResponse.DS.Tables[0].Rows[0]["FileLocation"]);
                                            string company_no = Convert.ToString(objCMDsResponse.DS.Tables[0].Rows[0]["CompanyNumber"]);
                                            string customerNumber = "";

                                            attachmentPath = attachmentPath + @"\" + locationCode + @"\" + filePath;

                                            if (!System.IO.Directory.Exists(tempfilePath))
                                                System.IO.Directory.CreateDirectory(tempfilePath);


                                            string attname = String.Format("{0}\\{1}", tempfilePath, att.Name);
                                            att.SaveAs(attname, true);

                                            DataSet dsExcel = new DataSet();
                                            if (extn.ToLower().Contains(".csv"))
                                            {
                                                // dsExcel = clsExcelHelper.ImportCSV(tempfilePath + @"\" + strFileName, true, ",", 0);
                                                // dsExcel = clsExcelHelper.ImportCSVNew(tempfilePath + @"\" + strFileName);
                                                dsExcel = clsExcelHelper.ConvertCSVtoDataSet(tempfilePath + @"\" + fileName);
                                            }
                                            else if (extn.ToLower().Contains(".xlsx"))
                                            {
                                                dsExcel = clsExcelHelper.ImportExcelXLSX(tempfilePath + @"\" + fileName, false);
                                            }
                                            else if (extn.ToLower().Contains(".xls"))
                                            {
                                                dsExcel = clsExcelHelper.ImportExcelXLS(tempfilePath + @"\" + fileName, false);
                                            }
                                            else if (extn.ToLower().Contains(".bbb"))
                                            {
                                                //  string strBillingCustomerNumber = ""; // = objCommon.GetConfigValue("BBBBillingCustomerNumber");
                                                //  string strServiceType = objCommon.GetConfigValue("BBBServiceType");
                                                //  string strEntredBy = objCommon.GetConfigValue("BBBEnteredBy");
                                                //  string strPickupDeliveryTransferFlag = objCommon.GetConfigValue("BBBPickupDeliveryTransferFlag");
                                                //  string strRequestedBy = objCommon.GetConfigValue("BBBRequestedBy");
                                                DataTable dtConfiguredData = new DataTable();
                                                clsCommon.DSResponse objDsRes = new clsCommon.DSResponse();
                                                objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);
                                                if (objDsRes.dsResp.ResponseVal)
                                                {
                                                    dtConfiguredData = objDsRes.DS.Tables[0];
                                                    customerNumber = Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                                                }
                                                else
                                                {
                                                    executionLogMessage = "Order Post Template Mapping Missing " + System.Environment.NewLine;
                                                    executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                                    executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                                    executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                                    executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                                    emailSubject = "Order Post Template Mapping Missing";
                                                    objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                                    objCommon.WriteExecutionLog(executionLogMessage);
                                                    continue;
                                                }
                                                if (productSubCode == "INB")
                                                {
                                                    dsExcel = ConvertBBBINBFlatfileToDataTable(tempfilePath + @"\" + fileName,
                                                         dtConfiguredData);
                                                }
                                                else if (productSubCode == "TND")
                                                {
                                                    dsExcel = ConvertBBBTNDFlatfileToDataTable(tempfilePath + @"\" + fileName,
                                                         dtConfiguredData);
                                                }
                                            }

                                            if (customerName == objCommon.GetConfigValue("TGTCustomerName") || customerName == objCommon.GetConfigValue("LBRANDSCustomerName"))
                                            {
                                                DataTable dtConfiguredData = new DataTable();
                                                clsCommon.DSResponse objDsRes = new clsCommon.DSResponse();
                                                objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);
                                                if (objDsRes.dsResp.ResponseVal)
                                                {
                                                    dtConfiguredData = objDsRes.DS.Tables[0];
                                                    customerNumber = Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                                                }
                                                else
                                                {
                                                    executionLogMessage = "Order Post Template Mapping Missing " + System.Environment.NewLine;
                                                    executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                                    executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                                    executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                                    executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                                    emailSubject = "Order Post Template Mapping Missing";
                                                    objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                                    objCommon.WriteExecutionLog(executionLogMessage);
                                                    continue;
                                                }
                                                if (customerName == objCommon.GetConfigValue("TGTCustomerName"))
                                                {
                                                    dsExcel = TGTGenerateOrderDataTable(dsExcel.Tables[0], dtConfiguredData, productSubCode);
                                                }
                                                else if (customerName == objCommon.GetConfigValue("LBRANDSCustomerName"))
                                                {
                                                    dsExcel = LASGenerateOrderDataTable(dsExcel.Tables[0], dtConfiguredData, productSubCode);
                                                }
                                            }
                                            else if (customerName == objCommon.GetConfigValue("CDSCustomerName"))
                                            {
                                                DataTable dtConfiguredData = new DataTable();
                                                clsCommon.DSResponse objDsRes = new clsCommon.DSResponse();
                                                objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);
                                                if (objDsRes.dsResp.ResponseVal)
                                                {
                                                    dtConfiguredData = objDsRes.DS.Tables[0];
                                                    customerNumber = Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                                                }
                                                else
                                                {
                                                    executionLogMessage = "Order Post Template Mapping Missing " + System.Environment.NewLine;
                                                    executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                                    executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                                    executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                                    executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                                    emailSubject = "Order Post Template Mapping Missing";
                                                    objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                                    objCommon.WriteExecutionLog(executionLogMessage);
                                                    continue;
                                                }
                                                if (customerName == objCommon.GetConfigValue("CDSCustomerName"))
                                                {
                                                    dsExcel = CDSGenerateOrderDataTable(dsExcel.Tables[0], dtConfiguredData, productSubCode);
                                                }
                                            }
                                            if (dsExcel.Tables.Count > 0)
                                            {
                                                DataTable dataTable = dsExcel.Tables[0];

                                                if (dataTable.Rows.Count < 1)
                                                {
                                                    executionLogMessage = "No data found in the file  :  " + attname + System.Environment.NewLine;
                                                    executionLogMessage += "But Still processed this file";
                                                    objCommon.WriteExecutionLog(executionLogMessage);
                                                }

                                                clsCommon.DSResponse objDsResponse = new clsCommon.DSResponse();

                                                // case switch case 
                                                // switch with integer type
                                                switch (productCode)
                                                {
                                                    case "FM":

                                                        if (customerName.ToUpper() == "BLUDOT" && dataTable.Columns.Count == 2)
                                                        {
                                                            DataTable dtable = new DataTable();
                                                            dtable.Clear();
                                                            dtable.Columns.Add("Customer_Reference");
                                                            dtable.Columns.Add("Service Type");
                                                            dtable.Columns.Add("Delivery Name");
                                                            dtable.Columns.Add("Delivery Address");
                                                            dtable.Columns.Add("Delivery State");
                                                            dtable.Columns.Add("Delivery City");
                                                            dtable.Columns.Add("Delivery Zip");
                                                            dtable.Columns.Add("Delivery Phone Number");
                                                            dtable.Columns.Add("Item Number");
                                                            dtable.Columns.Add("Item Description");
                                                            dtable.Columns.Add("Pieces");
                                                            dtable.Columns.Add("Weight");
                                                            dtable.Columns.Add("Return");
                                                            //dtable.Columns.Add("Bol_Number");

                                                            DataView view = new DataView(dataTable);
                                                            DataTable dtdistinctValues = view.ToTable(true, "Customer_Reference");

                                                            foreach (DataRow dr in dtdistinctValues.Rows)
                                                            {
                                                                DataRow[] drresult = dataTable.Select("Customer_Reference= '" + dr["Customer_Reference"] + "'");

                                                                object value = dr["Customer_Reference"];
                                                                if (value == DBNull.Value)
                                                                    break;
                                                                string strCustomer_Reference = Convert.ToString(dr["Customer_Reference"]);

                                                                //if (dr["Customer_Reference"] == DBNull.Value
                                                                //  && dr["Return"] == DBNull.Value)
                                                                //    dr.Delete();
                                                                //string strCustomer_Reference = Convert.ToString(dr["Customer_Reference"]);
                                                                string strReturn = Convert.ToString(drresult[0]["Return"]);

                                                                objDsResponse = objCommon.GetRouteStopDetails(strCustomer_Reference, company_no);
                                                                if (objDsResponse.dsResp.ResponseVal)
                                                                {
                                                                    string strInputFilePath = "";
                                                                    string strDatetime = DateTime.Now.ToString("yyyyMMddHHmmss");

                                                                    clsRoute objclsRoute = new clsRoute();
                                                                    clsCommon.ReturnResponse objresponse = new clsCommon.ReturnResponse();
                                                                    string uniqueid = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["unique_id_no"]);
                                                                    string GeneratedUniqueId = objclsRoute.GenerateUniqueNumber(Convert.ToInt32(company_no), Convert.ToInt32(uniqueid));

                                                                    objresponse = objclsRoute.DataTracRouteStopGetAPI(GeneratedUniqueId);
                                                                    // objresponse.ResponseVal = true;
                                                                    //  objresponse.Reason = "{\"00100035009\":{\"room\":null,\"unique_id\":35009,\"c2_paperwork\":false,\"company_number_text\":\"EASTRN TIME ON CENTRAL SERVER\",\"company_number\":1,\"addl_charge_code11\":null,\"billing_override_amt\":null,\"addl_charge_occur1\":null,\"updated_time\":null,\"stop_sequence\":\"0010\",\"phone\":null,\"city\":\"Alpharetta\",\"created_by\":\"DX*\",\"signature_images\":[],\"pricing_zone\":null,\"signature_filename\":null,\"addl_charge_code10\":null,\"cod_check_no\":null,\"length\":null,\"expected_weight\":null,\"actual_settlement_amt\":null,\"actual_pieces\":null,\"updated_date\":null,\"schedule_stop_id\":null,\"photos_exist\":false,\"stop_type_text\":\"Delivery\",\"stop_type\":\"D\",\"return\":false,\"addl_charge_code6\":null,\"dispatch_zone\":null,\"upload_time\":null,\"actual_cod_amt\":null,\"location_accuracy\":null,\"progress\":[{\"status_time\":\"10:22:02\",\"status_text\":\"Entered in carrier's system\",\"status_date\":\"2021-08-04\"}],\"received_route\":null,\"override_settle_percent\":null,\"cod_amount\":null,\"addl_charge_code9\":null,\"eta_date\":null,\"cod_type_text\":\"None\",\"cod_type\":\"0\",\"addl_charge_occur3\":null,\"reference\":null,\"sent_to_phone\":false,\"addl_charge_occur12\":null,\"callback_required_text\":\"No\",\"callback_required\":\"N\",\"service_level_text\":\"1HR RUSH SERVICE\",\"service_level\":6,\"original_id\":null,\"width\":null,\"received_sequence\":null,\"transfer_to_sequence\":null,\"cases\":null,\"times_sent\":0,\"transfer_to_route\":null,\"zip_code\":null,\"settlement_override_amt\":null,\"driver_app_status_text\":\"\",\"driver_app_status\":\"0\",\"route_code_text\":\"NAPA\",\"route_code\":\"NAPA\",\"received_shift\":null,\"addl_charge_occur6\":null,\"addl_charge_occur11\":null,\"vehicle\":null,\"addl_charge_code5\":null,\"addl_charge_occur9\":null,\"eta\":null,\"departure_time\":null,\"combine_data\":null,\"actual_latitude\":null,\"posted_by\":null,\"insurance_value\":null,\"return_redel_id\":null,\"addl_charge_code1\":null,\"origin_code_text\":\"Added using API\",\"origin_code\":\"A\",\"ordered_by\":null,\"posted_date\":null,\"actual_billing_amt\":null,\"created_date\":\"2021-08-04\",\"latitude\":null,\"received_pieces\":null,\"addl_charge_code7\":null,\"totes\":null,\"asn_sent\":0,\"comments\":null,\"verification_id_type_text\":\"None\",\"verification_id_type\":\"0\",\"posted_time\":null,\"item_scans_required\":true,\"shift_id\":null,\"addon_billing_amt\":null,\"actual_delivery_date\":null,\"id\":\"00100035009\",\"actual_arrival_time\":null,\"signature_required\":true,\"longitude\":null,\"expected_pieces\":null,\"loaded_pieces\":null,\"alt_lookup\":null,\"customer_number_text\":\"Routing Customer\",\"customer_number\":4999,\"created_time\":\"10:22:02\",\"addl_charge_code8\":null,\"signature\":null,\"actual_depart_time\":null,\"bol_number\":null,\"actual_cod_type_text\":\"None\",\"actual_cod_type\":\"0\",\"invoice_number\":null,\"branch_id\":null,\"special_instructions2\":null,\"updated_by\":null,\"verification_id_details\":null,\"required_signature_type_text\":\"Any signature\",\"required_signature_type\":\"0\",\"addl_charge_occur7\":null,\"orig_order_number\":null,\"special_instructions1\":null,\"notes\":[],\"image_sign_req\":true,\"attention\":null,\"minutes_late\":0,\"late_notice_time\":null,\"received_unique_id\":null,\"exception_code\":null,\"addl_charge_code4\":null,\"addl_charge_occur4\":null,\"redelivery\":false,\"addl_charge_occur10\":null,\"upload_date\":null,\"special_instructions4\":null,\"address_name\":null,\"addl_charge_occur8\":null,\"address_point_customer\":null,\"received_branch\":null,\"items\":[],\"return_redelivery_date\":null,\"height\":null,\"actual_longitude\":null,\"service_time\":null,\"phone_ext\":null,\"addl_charge_occur2\":null,\"late_notice_date\":null,\"address\":\"123 Stop Address Street\",\"arrival_time\":null,\"posted_status\":false,\"route_date\":\"2021-08-03\",\"addl_charge_code12\":null,\"addl_charge_code3\":null,\"return_redelivery_flag_text\":\"None\",\"return_redelivery_flag\":\"N\",\"additional_instructions\":null,\"updated_by_scanner\":false,\"special_instructions3\":null,\"addl_charge_occur5\":null,\"address_point\":0,\"actual_weight\":null,\"received_company\":null,\"addl_charge_code2\":null,\"state\":\"GA\"}}";
                                                                    // objresponse.Reason = "{\"00204352124\": {\"posted_by\": null, \"addon_billing_amt\": null, \"minutes_late\": 0, \"insurance_value\": null, \"addl_charge_occur5\": null, \"actual_pieces\": null, \"actual_depart_time\": null, \"created_time\": \"08:43:34\", \"cod_amount\": null, \"special_instructions3\": null, \"width\": null, \"ordered_by\": null, \"addl_charge_code1\": null, \"signature_filename\": null, \"updated_date\": null, \"latitude\": null, \"signature\": null, \"received_branch\": null, \"late_notice_time\": null, \"route_code_text\": \"HDHOLD\", \"route_code\": \"HDHOLD\", \"phone_ext\": null, \"addl_charge_occur1\": null, \"received_sequence\": null, \"address_name\": \"TEST1\", \"address_point_customer\": null, \"actual_cod_amt\": null, \"signature_required\": true, \"stop_type_text\": \"Delivery\", \"stop_type\": \"D\", \"origin_code_text\": \"Added using API\", \"origin_code\": \"A\", \"invoice_number\": null, \"addl_charge_code11\": null, \"addl_charge_code12\": null, \"length\": null, \"vehicle\": null, \"item_scans_required\": true, \"updated_by_scanner\": false, \"addl_charge_code10\": null, \"unique_id\": 4352124, \"attention\": null, \"items\": [{\"item_number\": \"item1\", \"item_description\": \"first item\", \"reference\": \"1\", \"rma_route\": null, \"upload_time\": null, \"rma_stop_id\": 0, \"width\": null, \"redelivery\": false, \"received_pieces\": null, \"cod_amount\": null, \"height\": null, \"comments\": null, \"actual_pieces\": null, \"actual_cod_amount\": null, \"rma_number\": null, \"manually_updated\": 0, \"unique_id\": 4352124, \"cod_type_text\": \"None\", \"cod_type\": \"0\", \"barcodes_unique\": false, \"actual_cod_type_text\": \"None\", \"actual_cod_type\": \"0\", \"return_redel_seq\": 0, \"expected_pieces\": 3, \"signature\": null, \"exception_code\": null, \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"updated_date\": null, \"expected_weight\": 50, \"created_date\": \"2021-08-26\", \"rma_origin\": null, \"created_by\": \"DX*\", \"loaded_pieces\": null, \"return_redelivery_flag_text\": \"\", \"return_redelivery_flag\": null, \"original_id\": 0, \"container_id\": \"container1\", \"return\": false, \"length\": null, \"notes\": [], \"actual_weight\": null, \"updated_by\": null, \"photos_exist\": false, \"second_container_id\": null, \"return_redel_id\": 0, \"asn_sent\": 0, \"actual_departure_time\": null, \"updated_time\": null, \"return_redelivery_date\": null, \"actual_arrival_time\": null, \"item_sequence\": 1, \"pallet_number\": null, \"actual_date\": null, \"insurance_value\": null, \"created_time\": \"08:43:34\", \"upload_date\": null, \"scans\": [], \"id\": \"002043521240001\", \"truck_id\": 0}, {\"item_number\": \"item2\", \"item_description\": \"first item\", \"reference\": \"1\", \"rma_route\": null, \"upload_time\": null, \"rma_stop_id\": 0, \"width\": null, \"redelivery\": false, \"received_pieces\": null, \"cod_amount\": null, \"height\": null, \"comments\": null, \"actual_pieces\": null, \"actual_cod_amount\": null, \"rma_number\": null, \"manually_updated\": 0, \"unique_id\": 4352124, \"cod_type_text\": \"None\", \"cod_type\": \"0\", \"barcodes_unique\": false, \"actual_cod_type_text\": \"None\", \"actual_cod_type\": \"0\", \"return_redel_seq\": 0, \"expected_pieces\": 1, \"signature\": null, \"exception_code\": null, \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"updated_date\": null, \"expected_weight\": 150, \"created_date\": \"2021-08-26\", \"rma_origin\": null, \"created_by\": \"DX*\", \"loaded_pieces\": null, \"return_redelivery_flag_text\": \"\", \"return_redelivery_flag\": null, \"original_id\": 0, \"container_id\": \"container2\", \"return\": false, \"length\": null, \"notes\": [], \"actual_weight\": null, \"updated_by\": null, \"photos_exist\": false, \"second_container_id\": null, \"return_redel_id\": 0, \"asn_sent\": 0, \"actual_departure_time\": null, \"updated_time\": null, \"return_redelivery_date\": null, \"actual_arrival_time\": null, \"item_sequence\": 2, \"pallet_number\": null, \"actual_date\": null, \"insurance_value\": null, \"created_time\": \"08:43:34\", \"upload_date\": null, \"scans\": [], \"id\": \"002043521240002\", \"truck_id\": 0}, {\"item_number\": \"item3\", \"item_description\": \"first item\", \"reference\": \"1\", \"rma_route\": null, \"upload_time\": null, \"rma_stop_id\": 0, \"width\": null, \"redelivery\": false, \"received_pieces\": null, \"cod_amount\": null, \"height\": null, \"comments\": null, \"actual_pieces\": null, \"actual_cod_amount\": null, \"rma_number\": null, \"manually_updated\": 0, \"unique_id\": 4352124, \"cod_type_text\": \"None\", \"cod_type\": \"0\", \"barcodes_unique\": false, \"actual_cod_type_text\": \"None\", \"actual_cod_type\": \"0\", \"return_redel_seq\": 0, \"expected_pieces\": 1, \"signature\": null, \"exception_code\": null, \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"updated_date\": null, \"expected_weight\": 250, \"created_date\": \"2021-08-26\", \"rma_origin\": null, \"created_by\": \"DX*\", \"loaded_pieces\": null, \"return_redelivery_flag_text\": \"\", \"return_redelivery_flag\": null, \"original_id\": 0, \"container_id\": \"container3\", \"return\": false, \"length\": null, \"notes\": [], \"actual_weight\": null, \"updated_by\": null, \"photos_exist\": false, \"second_container_id\": null, \"return_redel_id\": 0, \"asn_sent\": 0, \"actual_departure_time\": null, \"updated_time\": null, \"return_redelivery_date\": null, \"actual_arrival_time\": null, \"item_sequence\": 3, \"pallet_number\": null, \"actual_date\": null, \"insurance_value\": null, \"created_time\": \"08:43:35\", \"upload_date\": null, \"scans\": [], \"id\": \"002043521240003\", \"truck_id\": 0}, {\"item_number\": \"item4\", \"item_description\": \"first item\", \"reference\": \"1\", \"rma_route\": null, \"upload_time\": null, \"rma_stop_id\": 0, \"width\": null, \"redelivery\": false, \"received_pieces\": null, \"cod_amount\": null, \"height\": null, \"comments\": null, \"actual_pieces\": null, \"actual_cod_amount\": null, \"rma_number\": null, \"manually_updated\": 0, \"unique_id\": 4352124, \"cod_type_text\": \"None\", \"cod_type\": \"0\", \"barcodes_unique\": false, \"actual_cod_type_text\": \"None\", \"actual_cod_type\": \"0\", \"return_redel_seq\": 0, \"expected_pieces\": 21, \"signature\": null, \"exception_code\": null, \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"updated_date\": null, \"expected_weight\": 350, \"created_date\": \"2021-08-26\", \"rma_origin\": null, \"created_by\": \"DX*\", \"loaded_pieces\": null, \"return_redelivery_flag_text\": \"\", \"return_redelivery_flag\": null, \"original_id\": 0, \"container_id\": \"container4\", \"return\": false, \"length\": null, \"notes\": [], \"actual_weight\": null, \"updated_by\": null, \"photos_exist\": false, \"second_container_id\": null, \"return_redel_id\": 0, \"asn_sent\": 0, \"actual_departure_time\": null, \"updated_time\": null, \"return_redelivery_date\": null, \"actual_arrival_time\": null, \"item_sequence\": 4, \"pallet_number\": null, \"actual_date\": null, \"insurance_value\": null, \"created_time\": \"08:43:36\", \"upload_date\": null, \"scans\": [], \"id\": \"002043521240004\", \"truck_id\": 0}], \"addl_charge_occur10\": null, \"verification_id_type_text\": \"None\", \"verification_id_type\": \"0\", \"addl_charge_occur7\": null, \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"posted_time\": null, \"c2_paperwork\": false, \"original_id\": null, \"progress\": [{\"status_time\": \"08:43:34\", \"status_date\": \"2021-08-26\", \"status_text\": \"Entered in carrier's system\"}], \"service_level_text\": \"Basic Delivery\", \"service_level\": 56, \"created_by\": \"DX*\", \"required_signature_type_text\": \"Any signature\", \"required_signature_type\": \"0\", \"special_instructions1\": null, \"actual_billing_amt\": null, \"branch_id_text\": \"JWL Baltimore, MD\", \"branch_id\": \"BWI\", \"actual_cod_type_text\": \"None\", \"actual_cod_type\": \"0\", \"pricing_zone\": null, \"state\": \"TX\", \"signature_images\": [], \"special_instructions4\": null, \"photos_exist\": false, \"height\": null, \"eta_date\": null, \"upload_date\": null, \"zip_code\": \"75034\", \"actual_latitude\": null, \"override_settle_percent\": null, \"notes\": [{\"entry_time\": \"08:43:34\", \"note_text\": \"** Expected pieces: 0 -> 3\", \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"item_sequence\": null, \"entry_date\": \"2021-08-26\", \"user_entered\": false, \"show_to_cust\": false, \"note_type_text\": \"Stop\", \"note_type\": \"0\", \"unique_id\": 4352124, \"id\": \"00204352124    0020210826084334DX* 24\", \"user_id\": \"DX*\"}, {\"entry_time\": \"08:43:34\", \"note_text\": \"** Expected weight:      0 ->      50\", \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"item_sequence\": null, \"entry_date\": \"2021-08-26\", \"user_entered\": false, \"show_to_cust\": false, \"note_type_text\": \"Stop\", \"note_type\": \"0\", \"unique_id\": 4352124, \"id\": \"00204352124    0020210826084334DX* 25\", \"user_id\": \"DX*\"}, {\"entry_time\": \"08:43:35\", \"note_text\": \"** Expected pieces: 3 -> 4\", \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"item_sequence\": null, \"entry_date\": \"2021-08-26\", \"user_entered\": false, \"show_to_cust\": false, \"note_type_text\": \"Stop\", \"note_type\": \"0\", \"unique_id\": 4352124, \"id\": \"00204352124    0020210826084335DX* 24\", \"user_id\": \"DX*\"}, {\"entry_time\": \"08:43:35\", \"note_text\": \"** Expected weight:     50 ->     200\", \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"item_sequence\": null, \"entry_date\": \"2021-08-26\", \"user_entered\": false, \"show_to_cust\": false, \"note_type_text\": \"Stop\", \"note_type\": \"0\", \"unique_id\": 4352124, \"id\": \"00204352124    0020210826084335DX* 25\", \"user_id\": \"DX*\"}, {\"entry_time\": \"08:43:36\", \"note_text\": \"** Expected pieces: 4 -> 5\", \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"item_sequence\": null, \"entry_date\": \"2021-08-26\", \"user_entered\": false, \"show_to_cust\": false, \"note_type_text\": \"Stop\", \"note_type\": \"0\", \"unique_id\": 4352124, \"id\": \"00204352124    0020210826084336DX* 24\", \"user_id\": \"DX*\"}, {\"entry_time\": \"08:43:36\", \"note_text\": \"** Expected weight:    200 ->     450\", \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"item_sequence\": null, \"entry_date\": \"2021-08-26\", \"user_entered\": false, \"show_to_cust\": false, \"note_type_text\": \"Stop\", \"note_type\": \"0\", \"unique_id\": 4352124, \"id\": \"00204352124    0020210826084336DX* 25\", \"user_id\": \"DX*\"}, {\"entry_time\": \"08:43:37\", \"note_text\": \"** Expected pieces: 5 -> 26\", \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"item_sequence\": null, \"entry_date\": \"2021-08-26\", \"user_entered\": false, \"show_to_cust\": false, \"note_type_text\": \"Stop\", \"note_type\": \"0\", \"unique_id\": 4352124, \"id\": \"00204352124    0020210826084337DX* 24\", \"user_id\": \"DX*\"}, {\"entry_time\": \"08:43:37\", \"note_text\": \"** Expected weight:    450 ->     800\", \"company_number_text\": \"JW LOGISTICS EAST REGION\", \"company_number\": 2, \"item_sequence\": null, \"entry_date\": \"2021-08-26\", \"user_entered\": false, \"show_to_cust\": false, \"note_type_text\": \"Stop\", \"note_type\": \"0\", \"unique_id\": 4352124, \"id\": \"00204352124    0020210826084337DX* 25\", \"user_id\": \"DX*\"}], \"additional_instructions\": null, \"addl_charge_occur6\": null, \"driver_app_status_text\": \"\", \"driver_app_status\": \"0\", \"combine_data\": null, \"addl_charge_code2\": null, \"service_time\": null, \"city\": \"FRISCO\", \"room\": null, \"addl_charge_code7\": null, \"billing_override_amt\": null, \"totes\": null, \"sent_to_phone\": false, \"address\": \"1000 PARKWOOD BLVD\", \"posted_date\": null, \"phone\": \"111-111-1111\", \"late_notice_date\": null, \"received_route\": null, \"bol_number\": \"1\", \"asn_sent\": 0, \"addl_charge_occur3\": null, \"departure_time\": null, \"received_unique_id\": null, \"orig_order_number\": null, \"reference\": \"1\", \"comments\": null, \"updated_by\": null, \"customer_number_text\": \"HD - BWI 21229\", \"customer_number\": 516, \"addl_charge_code4\": null, \"addl_charge_code9\": null, \"location_accuracy\": null, \"verification_id_details\": null, \"cases\": null, \"actual_arrival_time\": null, \"received_company\": null, \"addl_charge_code5\": null, \"addl_charge_occur11\": null, \"addl_charge_code6\": null, \"actual_settlement_amt\": null, \"addl_charge_occur12\": null, \"cod_check_no\": null, \"updated_time\": null, \"expected_pieces\": 26, \"times_sent\": 0, \"addl_charge_occur9\": null, \"id\": \"00204352124\", \"route_date\": \"2021-08-31\", \"schedule_stop_id\": null, \"return\": false, \"addl_charge_occur4\": null, \"image_sign_req\": false, \"created_date\": \"2021-08-26\", \"longitude\": null, \"redelivery\": false, \"actual_weight\": null, \"cod_type_text\": \"None\", \"cod_type\": \"0\", \"eta\": null, \"transfer_to_sequence\": null, \"callback_required_text\": \"No\", \"callback_required\": \"N\", \"alt_lookup\": null, \"addl_charge_occur8\": null, \"posted_status\": false, \"addl_charge_occur2\": null, \"transfer_to_route\": null, \"shift_id\": null, \"addl_charge_code8\": null, \"upload_time\": null, \"received_shift\": null, \"return_redel_id\": null, \"addl_charge_code3\": null, \"stop_sequence\": \"0010\", \"dispatch_zone\": null, \"expected_weight\": 800, \"special_instructions2\": null, \"actual_longitude\": null, \"settlement_override_amt\": null, \"actual_delivery_date\": null, \"arrival_time\": null, \"return_redelivery_flag_text\": \"None\", \"return_redelivery_flag\": \"N\", \"loaded_pieces\": null, \"exception_code\": null, \"address_point\": 0, \"return_redelivery_date\": null, \"received_pieces\": null, \"_utc_offset\": \"-04:00\"}}";
                                                                    if (objresponse.ResponseVal)
                                                                    {
                                                                        executionLogMessage = "RouteStopGetAPI Success " + System.Environment.NewLine;
                                                                        // strExecutionLogMessage += "Request -" + request + System.Environment.NewLine;
                                                                        executionLogMessage += "Response -" + objresponse.Reason + System.Environment.NewLine;
                                                                        objCommon.WriteExecutionLog(executionLogMessage);
                                                                        // DataSet dsOrderResponse = objCommon.jsonToDataSet(objresponse.Reason, "RouteStopPostAPI");
                                                                        DataSet dsResponse = objCommon.jsonToDataSet(objresponse.Reason, "RouteStopPutAPI");
                                                                        var UniqueId = Convert.ToString(dsResponse.Tables[0].Rows[0]["id"]);
                                                                        try
                                                                        {
                                                                            if (dsResponse.Tables.Contains("items"))
                                                                            {
                                                                                List<RouteStopResponseItem> itemList = new List<RouteStopResponseItem>();
                                                                                for (int i = 0; i < dsResponse.Tables["items"].Rows.Count; i++)
                                                                                {
                                                                                    RouteStopResponseItem item = new RouteStopResponseItem();
                                                                                    DataTable dt = new DataTable();
                                                                                    dt = dsResponse.Tables["items"];


                                                                                    DataRow _newRow = dtable.NewRow();

                                                                                    _newRow["Customer_Reference"] = strCustomer_Reference;
                                                                                    _newRow["Service Type"] = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["stop_signature"]);
                                                                                    _newRow["Delivery Name"] = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["stop_name"]);
                                                                                    _newRow["Delivery Address"] = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["stop_address"]);
                                                                                    _newRow["Delivery City"] = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["stop_city"]);
                                                                                    _newRow["Delivery State"] = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["stop_state"]);
                                                                                    _newRow["Delivery Zip"] = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["stop_zip_postal_code"]);
                                                                                    _newRow["Delivery Phone Number"] = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["stop_phone_no"]);
                                                                                    _newRow["Item Number"] = dt.Rows[i]["item_number"];
                                                                                    _newRow["Item Description"] = dt.Rows[i]["item_description"];
                                                                                    _newRow["Pieces"] = dt.Rows[i]["actual_pieces"];
                                                                                    _newRow["Weight"] = dt.Rows[i]["actual_weight"];
                                                                                    _newRow["Return"] = strReturn;
                                                                                    dtable.Rows.Add(_newRow);
                                                                                }
                                                                            }

                                                                        }
                                                                        catch (Exception ex)
                                                                        {
                                                                            executionLogMessage = "RouteStopGetFiles Exception -" + ex.Message + System.Environment.NewLine;
                                                                            executionLogMessage += "File Path is  -" + strInputFilePath + System.Environment.NewLine;
                                                                            executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                                            executionLogMessage += "For Reference -" + strCustomer_Reference + System.Environment.NewLine;
                                                                            //objCommon.WriteExecutionLog(strExecutionLogFileLocation, strExecutionLogMessage);
                                                                            objCommon.WriteErrorLog(ex, executionLogMessage);

                                                                            ErrorResponse objErrorResponse = new ErrorResponse();
                                                                            objErrorResponse.error = "Found exception while processing the record";
                                                                            objErrorResponse.status = "Error";
                                                                            objErrorResponse.code = "Excception while procesing the record.";
                                                                            objErrorResponse.reference = strCustomer_Reference;
                                                                            string strErrorResponse = JsonConvert.SerializeObject(objErrorResponse);
                                                                            DataSet dsFailureResponse = objCommon.jsonToDataSet(strErrorResponse);
                                                                            dsFailureResponse.Tables[0].TableName = "RouteStopGetAPI";
                                                                            objCommon.WriteDataToCsvFile(dsFailureResponse.Tables[0],
                                                           fileName, strDatetime);
                                                                            continue;
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            dtable.TableName = "Template";
                                                            IsSeen = true;
                                                            if (!info.Read && IsSeen)
                                                            {
                                                                oClient.MarkAsRead(info, true);
                                                                executionLogMessage = "Mark email as read, From :  " + oMail.From.ToString() + " , ReceivedDate" + oMail.ReceivedDate;
                                                                objCommon.WriteExecutionLog(executionLogMessage);
                                                            }
                                                            clsExcelHelper.ExportDataToXLSX(dtable, attachmentPath, fileName);
                                                            objCommon.CleanAttachmentWorkingFolder();
                                                        }
                                                        else
                                                        {
                                                            objDsResponse = objCommon.GetRouteStopPostTemplateDetails(customerName, locationCode, productCode);
                                                            if (objDsResponse.dsResp.ResponseVal)
                                                            {
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Customer_Reference"])].ColumnName = "Customer_Reference";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["ServiceType"])].ColumnName = "Service Type";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["DeliveryName"])].ColumnName = "Delivery Name";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["DeliveryAddress"])].ColumnName = "Delivery Address";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["DeliveryCity"])].ColumnName = "Delivery City";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["DeliveryState"])].ColumnName = "Delivery State";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["DeliveryZip"])].ColumnName = "Delivery Zip";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["DeliveryPhoneNumber"])].ColumnName = "Delivery Phone Number";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["ItemNumber"])].ColumnName = "Item Number";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["ItemDescription"])].ColumnName = "Item Description";
                                                                dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Pieces"])].ColumnName = "Pieces";
                                                                //  dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Weight"])].ColumnName = "Weight";
                                                                if (!string.IsNullOrEmpty(Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Weight"])))
                                                                {
                                                                    dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Weight"])].ColumnName = "Weight";
                                                                }
                                                                if (!string.IsNullOrEmpty(Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Return"])))
                                                                {
                                                                    dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Return"])].ColumnName = "Return";
                                                                }
                                                                if (!string.IsNullOrEmpty(Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Bol_Number"])))
                                                                {
                                                                    dataTable.Columns[Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["Bol_Number"])].ColumnName = "Bol Number";
                                                                }
                                                                dataTable.TableName = "Template";
                                                                IsSeen = true;
                                                                if (!info.Read && IsSeen)
                                                                {
                                                                    oClient.MarkAsRead(info, true);
                                                                    executionLogMessage = "Mark email as read, From :  " + oMail.From.ToString() + " , ReceivedDate" + oMail.ReceivedDate;
                                                                    objCommon.WriteExecutionLog(executionLogMessage);
                                                                }
                                                                clsExcelHelper.ExportDataToXLSX(dataTable, attachmentPath, fileName, customerName);
                                                                objCommon.CleanAttachmentWorkingFolder();

                                                            }
                                                            else
                                                            {
                                                                executionLogMessage = "RouteStop Post Template Mapping Missing " + System.Environment.NewLine;
                                                                executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                                                executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                                                executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                                                executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                                                emailSubject = "RouteStop Post Template Mapping Missing";
                                                                objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                                                objCommon.WriteExecutionLog(executionLogMessage);
                                                            }
                                                        }
                                                        break;

                                                    case "OD":

                                                        DataTable dtableOrderTemplate = new DataTable();
                                                        dtableOrderTemplate.Clear();

                                                        dtableOrderTemplate.Columns.Add("Delivery Date");
                                                        dtableOrderTemplate.Columns.Add("Company");
                                                        dtableOrderTemplate.Columns.Add("Billing Customer Number");
                                                        dtableOrderTemplate.Columns.Add("Customer Reference");
                                                        dtableOrderTemplate.Columns.Add("BOL Number");
                                                        dtableOrderTemplate.Columns.Add("Customer Name");
                                                        dtableOrderTemplate.Columns.Add("Route Number");
                                                        dtableOrderTemplate.Columns.Add("Original Driver No");
                                                        dtableOrderTemplate.Columns.Add("Correct Driver Number");
                                                        dtableOrderTemplate.Columns.Add("Carrier Name");
                                                        dtableOrderTemplate.Columns.Add("Address");
                                                        dtableOrderTemplate.Columns.Add("City");
                                                        dtableOrderTemplate.Columns.Add("State");
                                                        dtableOrderTemplate.Columns.Add("Zip");
                                                        dtableOrderTemplate.Columns.Add("Pieces");
                                                        dtableOrderTemplate.Columns.Add("Miles");
                                                        dtableOrderTemplate.Columns.Add("Delivery Zip");
                                                        dtableOrderTemplate.Columns.Add("Zip Code Surcharge?");
                                                        dtableOrderTemplate.Columns.Add("Store Code");
                                                        dtableOrderTemplate.Columns.Add("Type of Delivery");
                                                        dtableOrderTemplate.Columns.Add("Service Type");
                                                        dtableOrderTemplate.Columns.Add("Bill Rate");
                                                        dtableOrderTemplate.Columns.Add("Pieces ACC");
                                                        dtableOrderTemplate.Columns.Add("FSC");
                                                        dtableOrderTemplate.Columns.Add("Total Bill");
                                                        dtableOrderTemplate.Columns.Add("Carrier Base Pay");
                                                        dtableOrderTemplate.Columns.Add("Carrier ACC");
                                                        dtableOrderTemplate.Columns.Add("Carrier FSC");
                                                        dtableOrderTemplate.Columns.Add("Side Notes");
                                                        dtableOrderTemplate.Columns.Add("Pickup requested date");
                                                        dtableOrderTemplate.Columns.Add("Pickup will be ready by");
                                                        dtableOrderTemplate.Columns.Add("Pickup no later than");
                                                        dtableOrderTemplate.Columns.Add("Pickup actual date");
                                                        dtableOrderTemplate.Columns.Add("Pickup actual arrival time");
                                                        dtableOrderTemplate.Columns.Add("Pickup actual depart time");
                                                        dtableOrderTemplate.Columns.Add("Pickup name");
                                                        dtableOrderTemplate.Columns.Add("Pickup address");
                                                        dtableOrderTemplate.Columns.Add("Pickup city");
                                                        dtableOrderTemplate.Columns.Add("Pickup state/province");
                                                        dtableOrderTemplate.Columns.Add("Pickup zip/postal code");
                                                        dtableOrderTemplate.Columns.Add("Pickup text signature");
                                                        dtableOrderTemplate.Columns.Add("Delivery requested date");
                                                        dtableOrderTemplate.Columns.Add("Deliver no earlier than");
                                                        dtableOrderTemplate.Columns.Add("Deliver no later than");
                                                        dtableOrderTemplate.Columns.Add("Delivery actual date");
                                                        dtableOrderTemplate.Columns.Add("Delivery actual arrive time");
                                                        dtableOrderTemplate.Columns.Add("Delivery actual depart time");
                                                        dtableOrderTemplate.Columns.Add("Delivery text signature");
                                                        dtableOrderTemplate.Columns.Add("Requested by");
                                                        dtableOrderTemplate.Columns.Add("Entered by");
                                                        dtableOrderTemplate.Columns.Add("Pickup Delivery Transfer Flag");
                                                        dtableOrderTemplate.Columns.Add("Weight");
                                                        dtableOrderTemplate.Columns.Add("Insurance Amount");
                                                        dtableOrderTemplate.Columns.Add("Master airway bill number");
                                                        dtableOrderTemplate.Columns.Add("PO Number");
                                                        dtableOrderTemplate.Columns.Add("House airway bill number");
                                                        // dtableOrderTemplate.Columns.Add("Dimensions");
                                                        dtableOrderTemplate.Columns.Add("Item Number");
                                                        dtableOrderTemplate.Columns.Add("Item Description");
                                                        dtableOrderTemplate.Columns.Add("Dim Height");
                                                        dtableOrderTemplate.Columns.Add("Dim Length");
                                                        dtableOrderTemplate.Columns.Add("Dim Width");
                                                        dtableOrderTemplate.Columns.Add("Pickup Room");
                                                        dtableOrderTemplate.Columns.Add("Pickup Attention");
                                                        dtableOrderTemplate.Columns.Add("Deliver Attention");

                                                        //dtableOrderTemplate.Columns.Add("rate_buck_amt1");
                                                        dtableOrderTemplate.Columns.Add("rate_buck_amt2");
                                                        //  dtableOrderTemplate.Columns.Add("rate_buck_amt3");
                                                        dtableOrderTemplate.Columns.Add("rate_buck_amt4");
                                                        dtableOrderTemplate.Columns.Add("rate_buck_amt5");
                                                        dtableOrderTemplate.Columns.Add("rate_buck_amt6");
                                                        dtableOrderTemplate.Columns.Add("rate_buck_amt7");
                                                        dtableOrderTemplate.Columns.Add("rate_buck_amt8");
                                                        dtableOrderTemplate.Columns.Add("rate_buck_amt9");
                                                        //  dtableOrderTemplate.Columns.Add("rate_buck_amt10");
                                                        dtableOrderTemplate.Columns.Add("rate_buck_amt11");

                                                        //  dtableOrderTemplate.Columns.Add("charge1");
                                                        dtableOrderTemplate.Columns.Add("charge2");
                                                        dtableOrderTemplate.Columns.Add("charge3");
                                                        dtableOrderTemplate.Columns.Add("charge4");
                                                        // dtableOrderTemplate.Columns.Add("charge5");
                                                        //  dtableOrderTemplate.Columns.Add("charge6");

                                                        clsCommon.DSResponse objDsRes = new clsCommon.DSResponse();
                                                        objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);
                                                        if (objDsRes.dsResp.ResponseVal)
                                                        {
                                                            string strDatetime = DateTime.Now.ToString("yyyyMMddHHmmss");
                                                            DataTable dtOrderData = dsExcel.Tables[0];
                                                            GenerateOrderTemplate(ref dtableOrderTemplate, dtOrderData, fileName, customerName, locationCode, productCode, productSubCode);

                                                            //try
                                                            //{
                                                            //    DataTable dtOrderData = dsExcel.Tables[0];

                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_Date"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_Date"])].ColumnName = "Delivery Date";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Billing_Customer_Number"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Billing_Customer_Number"])].ColumnName = "Billing Customer Number";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Customer_Reference"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Customer_Reference"])].ColumnName = "Customer Reference";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["BOL_Number"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["BOL_Number"])].ColumnName = "BOL Number";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Customer_Name"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Customer_Name"])].ColumnName = "Customer Name";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Route_Number"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Route_Number"])].ColumnName = "Route Number";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Original_Driver_No"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Original_Driver_No"])].ColumnName = "Original Driver No";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Correct_Driver_Number"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Correct_Driver_Number"])].ColumnName = "Correct Driver Number";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_Name"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_Name"])].ColumnName = "Carrier Name";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Address"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Address"])].ColumnName = "Address";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["City"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["City"])].ColumnName = "City";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["State"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["State"])].ColumnName = "State";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Zip"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Zip"])].ColumnName = "Zip";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pieces"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pieces"])].ColumnName = "Pieces";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Miles"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Miles"])].ColumnName = "Miles";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_Zip"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_Zip"])].ColumnName = "Delivery Zip";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Zip_Code_Surcharge?"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Zip_Code_Surcharge?"])].ColumnName = "Zip Code Surcharge?";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Store_Code"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Store_Code"])].ColumnName = "Store_Code";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Type_of_Delivery"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Type_of_Delivery"])].ColumnName = "Type of Delivery";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Service_Type"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Service_Type"])].ColumnName = "Service Type";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Bill_Rate"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Bill_Rate"])].ColumnName = "Bill Rate";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pieces_ACC"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pieces_ACC"])].ColumnName = "Pieces ACC";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["FSC"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["FSC"])].ColumnName = "FSC";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Total_Bill"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Total_Bill"])].ColumnName = "Total Bill";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_Base_Pay"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_Base_Pay"])].ColumnName = "Carrier Base Pay";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_ACC"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_ACC"])].ColumnName = "Carrier ACC";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Side_Notes"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Side_Notes"])].ColumnName = "Side Notes";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_requested_date"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_requested_date"])].ColumnName = "Pickup requested date";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_will_be_ready_by"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_will_be_ready_by"])].ColumnName = "Pickup will be ready by";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_no_later_than"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_no_later_than"])].ColumnName = "Pickup no later than";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_date"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_date"])].ColumnName = "Pickup actual date";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_arrival_time"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_arrival_time"])].ColumnName = "Pickup actual arrival time";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_depart_time"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_depart_time"])].ColumnName = "Pickup actual depart time";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_name"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_name"])].ColumnName = "Pickup name";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_address"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_address"])].ColumnName = "Pickup address";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_city"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_city"])].ColumnName = "Pickup city";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_state/province"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_state/province"])].ColumnName = "Pickup state/province";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_zip/postal_code"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_zip/postal_code"])].ColumnName = "Pickup zip/postal code";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_text_signature"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_text_signature"])].ColumnName = "Pickup text signature";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_requested_date"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_requested_date"])].ColumnName = "Delivery requested date";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Deliver_no_earlier_than"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Deliver_no_earlier_than"])].ColumnName = "Deliver no earlier than";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Deliver_no_later_than"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Deliver_no_later_than"])].ColumnName = "Deliver no later than";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_date"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_date"])].ColumnName = "Delivery actual date";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_arrive_time"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_arrive_time"])].ColumnName = "Delivery actual arrive time";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_depart_time"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_depart_time"])].ColumnName = "Delivery actual depart time";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_text_signature"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_text_signature"])].ColumnName = "Delivery text signature";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Requested_by"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Requested_by"])].ColumnName = "Requested by";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Entered_by"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Entered_by"])].ColumnName = "Entered by";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_Delivery_Transfer_Flag"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_Delivery_Transfer_Flag"])].ColumnName = "Pickup Delivery Transfer Flag";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["weight"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["weight"])].ColumnName = "Weight";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["insurance_amount"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["insurance_amount"])].ColumnName = "Insurance Amount";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["master_airway_bill_number"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["master_airway_bill_number"])].ColumnName = "Master airway bill number";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["po_number"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["po_number"])].ColumnName = "PO Number";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["house_airway_bill_number"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["house_airway_bill_number"])].ColumnName = "House airway bill number";
                                                            //    }

                                                            //    //if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Dimensions"])))
                                                            //    //{
                                                            //    //    dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Dimensions"])].ColumnName = "Dimensions";
                                                            //    //}
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["item_number"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["item_number"])].ColumnName = "Item Number";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["item_description"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["item_description"])].ColumnName = "Item Description";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_height"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_height"])].ColumnName = "Dim Height";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_length"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_length"])].ColumnName = "Dim Length";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_width"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_width"])].ColumnName = "Dim Width";
                                                            //    }

                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["pickup_room"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["pickup_room"])].ColumnName = "Pickup Room";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["pickup_attention"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["pickup_attention"])].ColumnName = "Pickup Attention";
                                                            //    }
                                                            //    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["deliver_attention"])))
                                                            //    {
                                                            //        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["deliver_attention"])].ColumnName = "Deliver Attention";
                                                            //    }


                                                            //    if (dtOrderData.Rows.Count > 0)
                                                            //    {
                                                            //        foreach (DataRow dr in dtOrderData.Rows)
                                                            //        {
                                                            //            DataRow _newRow = dtableOrderTemplate.NewRow();
                                                            //            if (dr.Table.Columns.Contains("Delivery Date"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery Date"])))
                                                            //                {
                                                            //                    _newRow["Delivery Date"] = Convert.ToString(dr["Delivery Date"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Delivery Date"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Delivery Date"] = "";
                                                            //            }

                                                            //            _newRow["Company"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["CompanyNumber"]);

                                                            //            if (dr.Table.Columns.Contains("Billing Customer Number"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Billing Customer Number"])))
                                                            //                {
                                                            //                    _newRow["Billing Customer Number"] = Convert.ToString(dr["Billing Customer Number"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Billing Customer Number"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Billing Customer Number"] = "";
                                                            //            }

                                                            //            if (dr.Table.Columns.Contains("Customer Reference"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Customer Reference"])))
                                                            //                {
                                                            //                    _newRow["Customer Reference"] = Convert.ToString(dr["Customer Reference"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Customer Reference"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Customer Reference"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("BOL Number"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["BOL Number"])))
                                                            //                {
                                                            //                    _newRow["BOL Number"] = Convert.ToString(dr["BOL Number"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["BOL Number"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["BOL Number"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Customer Name"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Customer Name"])))
                                                            //                {
                                                            //                    _newRow["Customer Name"] = Convert.ToString(dr["Customer Name"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Customer Name"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Customer Name"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Route Number"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Route Number"])))
                                                            //                {
                                                            //                    _newRow["Route Number"] = Convert.ToString(dr["Route Number"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Route Number"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Route Number"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Original Driver No"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Original Driver No"])))
                                                            //                {
                                                            //                    _newRow["Original Driver No"] = Convert.ToString(dr["Original Driver No"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Original Driver No"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Original Driver No"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Correct Driver Number"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Correct Driver Number"])))
                                                            //                {
                                                            //                    _newRow["Correct Driver Number"] = Convert.ToString(dr["Correct Driver Number"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Correct Driver Number"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Correct Driver Number"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Carrier Name"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Carrier Name"])))
                                                            //                {
                                                            //                    _newRow["Carrier Name"] = Convert.ToString(dr["Carrier Name"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Carrier Name"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Carrier Name"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Address"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Address"])))
                                                            //                {
                                                            //                    _newRow["Address"] = Convert.ToString(dr["Address"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Address"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Address"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("City"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["City"])))
                                                            //                {
                                                            //                    _newRow["City"] = Convert.ToString(dr["City"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["City"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["City"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("State"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["State"])))
                                                            //                {
                                                            //                    _newRow["State"] = Convert.ToString(dr["State"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["State"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["State"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Zip"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Zip"])))
                                                            //                {
                                                            //                    _newRow["Zip"] = Convert.ToString(dr["Zip"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Zip"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Zip"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pieces"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                                                            //                {
                                                            //                    _newRow["Pieces"] = Convert.ToString(dr["Pieces"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pieces"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pieces"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Miles"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Miles"])))
                                                            //                {
                                                            //                    _newRow["Miles"] = Convert.ToString(dr["Miles"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Miles"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Miles"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Delivery Zip"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery Zip"])))
                                                            //                {
                                                            //                    _newRow["Delivery Zip"] = Convert.ToString(dr["Delivery Zip"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Delivery Zip"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Delivery Zip"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Zip Code Surcharge?"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Zip Code Surcharge?"])))
                                                            //                {
                                                            //                    _newRow["Zip Code Surcharge?"] = Convert.ToString(dr["Zip Code Surcharge?"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Zip Code Surcharge?"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Zip Code Surcharge?"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Store Code"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Store Code"])))
                                                            //                {
                                                            //                    _newRow["Store Code"] = Convert.ToString(dr["Store Code"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Store Code"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Store Code"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Type of Delivery"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Type of Delivery"])))
                                                            //                {
                                                            //                    _newRow["Type of Delivery"] = Convert.ToString(dr["Type of Delivery"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Type of Delivery"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Type of Delivery"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Service Type"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Service Type"])))
                                                            //                {
                                                            //                    _newRow["Service Type"] = Convert.ToString(dr["Service Type"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Service Type"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Service Type"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Bill Rate"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Bill Rate"])))
                                                            //                {
                                                            //                    _newRow["Bill Rate"] = Convert.ToString(dr["Bill Rate"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Bill Rate"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Bill Rate"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pieces ACC"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pieces ACC"])))
                                                            //                {
                                                            //                    _newRow["Pieces ACC"] = Convert.ToString(dr["Pieces ACC"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pieces ACC"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pieces ACC"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("FSC"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["FSC"])))
                                                            //                {
                                                            //                    _newRow["FSC"] = Convert.ToString(dr["FSC"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["FSC"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["FSC"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Total Bill"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Total Bill"])))
                                                            //                {
                                                            //                    _newRow["Total Bill"] = Convert.ToString(dr["Total Bill"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Total Bill"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Total Bill"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Carrier Base Pay"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Carrier Base Pay"])))
                                                            //                {
                                                            //                    _newRow["Carrier Base Pay"] = Convert.ToString(dr["Carrier Base Pay"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Carrier Base Pay"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Carrier Base Pay"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Carrier ACC"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Carrier ACC"])))
                                                            //                {
                                                            //                    _newRow["Carrier ACC"] = Convert.ToString(dr["Carrier ACC"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Carrier ACC"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Carrier ACC"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Side Notes"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Side Notes"])))
                                                            //                {
                                                            //                    _newRow["Side Notes"] = Convert.ToString(dr["Side Notes"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Side Notes"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Side Notes"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup requested date"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup requested date"])))
                                                            //                {
                                                            //                    _newRow["Pickup requested date"] = Convert.ToString(dr["Pickup requested date"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup requested date"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup requested date"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup will be ready by"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup will be ready by"])))
                                                            //                {
                                                            //                    _newRow["Pickup will be ready by"] = Convert.ToString(dr["Pickup will be ready by"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup will be ready by"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup will be ready by"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup no later than"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup no later than"])))
                                                            //                {
                                                            //                    _newRow["Pickup no later than"] = Convert.ToString(dr["Pickup no later than"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup no later than"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup no later than"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup actual date"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup actual date"])))
                                                            //                {
                                                            //                    _newRow["Pickup actual date"] = Convert.ToString(dr["Pickup actual date"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup actual date"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup actual date"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup actual arrival time"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup actual arrival time"])))
                                                            //                {
                                                            //                    _newRow["Pickup actual arrival time"] = Convert.ToString(dr["Pickup actual arrival time"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup actual arrival time"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup actual arrival time"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup actual depart time"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup actual depart time"])))
                                                            //                {
                                                            //                    _newRow["Pickup actual depart time"] = Convert.ToString(dr["Pickup actual depart time"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup actual depart time"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup actual depart time"] = "";
                                                            //            }

                                                            //            if (dr.Table.Columns.Contains("Pickup name"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup name"])))
                                                            //                {
                                                            //                    _newRow["Pickup name"] = Convert.ToString(dr["Pickup name"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup name"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup name"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup address"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup address"])))
                                                            //                {
                                                            //                    _newRow["Pickup address"] = Convert.ToString(dr["Pickup address"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup address"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup address"] = "";
                                                            //            }

                                                            //            if (dr.Table.Columns.Contains("Pickup city"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup city"])))
                                                            //                {
                                                            //                    _newRow["Pickup city"] = Convert.ToString(dr["Pickup city"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup city"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup city"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup state/province"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup state/province"])))
                                                            //                {
                                                            //                    _newRow["Pickup state/province"] = Convert.ToString(dr["Pickup state/province"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup state/province"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup state/province"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup zip/postal code"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup zip/postal code"])))
                                                            //                {
                                                            //                    _newRow["Pickup zip/postal code"] = Convert.ToString(dr["Pickup zip/postal code"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup zip/postal code"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup zip/postal code"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup text signature"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup text signature"])))
                                                            //                {
                                                            //                    _newRow["Pickup text signature"] = Convert.ToString(dr["Pickup text signature"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup text signature"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup text signature"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Delivery requested date"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery requested date"])))
                                                            //                {
                                                            //                    _newRow["Delivery requested date"] = Convert.ToString(dr["Delivery requested date"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Delivery requested date"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Delivery requested date"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Deliver no earlier than"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Deliver no earlier than"])))
                                                            //                {
                                                            //                    _newRow["Deliver no earlier than"] = Convert.ToString(dr["Deliver no earlier than"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Deliver no earlier than"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Deliver no earlier than"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Deliver no later than"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Deliver no later than"])))
                                                            //                {
                                                            //                    _newRow["Deliver no later than"] = Convert.ToString(dr["Deliver no later than"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Deliver no later than"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Deliver no later than"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Delivery actual date"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery actual date"])))
                                                            //                {
                                                            //                    _newRow["Delivery actual date"] = Convert.ToString(dr["Delivery actual date"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Delivery actual date"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Delivery actual date"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Delivery actual arrive time"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery actual arrive time"])))
                                                            //                {
                                                            //                    _newRow["Delivery actual arrive time"] = Convert.ToString(dr["Delivery actual arrive time"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Delivery actual arrive time"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Delivery actual arrive time"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Delivery actual depart time"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery actual depart time"])))
                                                            //                {
                                                            //                    _newRow["Delivery actual depart time"] = Convert.ToString(dr["Delivery actual depart time"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Delivery actual depart time"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Delivery actual depart time"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Delivery text signature"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery text signature"])))
                                                            //                {
                                                            //                    _newRow["Delivery text signature"] = Convert.ToString(dr["Delivery text signature"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Delivery text signature"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Delivery text signature"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Requested by"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Requested by"])))
                                                            //                {
                                                            //                    _newRow["Requested by"] = Convert.ToString(dr["Requested by"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Requested by"] = "";

                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Requested by"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Entered by"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Entered by"])))
                                                            //                {
                                                            //                    _newRow["Entered by"] = Convert.ToString(dr["Entered by"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Entered by"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Entered by"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup Delivery Transfer Flag"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup Delivery Transfer Flag"])))
                                                            //                {
                                                            //                    _newRow["Pickup Delivery Transfer Flag"] = Convert.ToString(dr["Pickup Delivery Transfer Flag"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup Delivery Transfer Flag"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup Delivery Transfer Flag"] = "";
                                                            //            }

                                                            //            if (dr.Table.Columns.Contains("Weight"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Weight"])))
                                                            //                {
                                                            //                    _newRow["Weight"] = Convert.ToString(dr["Weight"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Weight"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Weight"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Insurance Amount"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Insurance Amount"])))
                                                            //                {
                                                            //                    _newRow["Insurance Amount"] = Convert.ToString(dr["Insurance Amount"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Insurance Amount"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Insurance Amount"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Master airway bill number"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Master airway bill number"])))
                                                            //                {
                                                            //                    _newRow["Master airway bill number"] = Convert.ToString(dr["Master airway bill number"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Master airway bill number"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Master airway bill number"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("PO Number"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["PO Number"])))
                                                            //                {
                                                            //                    _newRow["PO Number"] = Convert.ToString(dr["PO Number"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["PO Number"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["PO Number"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("House airway bill number"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["House airway bill number"])))
                                                            //                {
                                                            //                    _newRow["House airway bill number"] = Convert.ToString(dr["House airway bill number"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["House airway bill number"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["House airway bill number"] = "";
                                                            //            }
                                                            //            //if (dr.Table.Columns.Contains("Dimensions"))
                                                            //            //{
                                                            //            //    if (!string.IsNullOrEmpty(Convert.ToString(dr["Dimensions"])))
                                                            //            //    {
                                                            //            //        _newRow["Dimensions"] = Convert.ToString(dr["Dimensions"]);
                                                            //            //    }
                                                            //            //    else
                                                            //            //    {
                                                            //            //        _newRow["Dimensions"] = "";
                                                            //            //    }
                                                            //            //}
                                                            //            //else
                                                            //            //{
                                                            //            //    _newRow["Dimensions"] = "";
                                                            //            //}

                                                            //            if (dr.Table.Columns.Contains("Item Number"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Item Number"])))
                                                            //                {
                                                            //                    _newRow["Item Number"] = Convert.ToString(dr["Item Number"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Item Number"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Item Number"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Item Description"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Item Description"])))
                                                            //                {
                                                            //                    _newRow["Item Description"] = Convert.ToString(dr["Item Description"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Item Description"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Item Description"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Dim Height"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Dim Height"])))
                                                            //                {
                                                            //                    _newRow["Dim Height"] = Convert.ToString(dr["Dim Height"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Dim Height"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Dim Height"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Dim Length"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Dim Length"])))
                                                            //                {
                                                            //                    _newRow["Dim Length"] = Convert.ToString(dr["Dim Length"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Dim Length"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Dim Length"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Dim Width"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Dim Width"])))
                                                            //                {
                                                            //                    _newRow["Dim Width"] = Convert.ToString(dr["Dim Width"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Dim Width"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Dim Width"] = "";
                                                            //            }

                                                            //            if (dr.Table.Columns.Contains("Pickup Room"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup Room"])))
                                                            //                {
                                                            //                    _newRow["Pickup Room"] = Convert.ToString(dr["Pickup Room"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup Room"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup Room"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Pickup Attention"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup Attention"])))
                                                            //                {
                                                            //                    _newRow["Pickup Attention"] = Convert.ToString(dr["Pickup Attention"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Pickup Attention"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Pickup Attention"] = "";
                                                            //            }
                                                            //            if (dr.Table.Columns.Contains("Deliver Attention"))
                                                            //            {
                                                            //                if (!string.IsNullOrEmpty(Convert.ToString(dr["Deliver Attention"])))
                                                            //                {
                                                            //                    _newRow["Deliver Attention"] = Convert.ToString(dr["Deliver Attention"]);
                                                            //                }
                                                            //                else
                                                            //                {
                                                            //                    _newRow["Deliver Attention"] = "";
                                                            //                }
                                                            //            }
                                                            //            else
                                                            //            {
                                                            //                _newRow["Deliver Attention"] = "";
                                                            //            }
                                                            //            dtableOrderTemplate.Rows.Add(_newRow);
                                                            //        }
                                                            //    }
                                                            //}
                                                            //catch (Exception ex)
                                                            //{
                                                            //    strExecutionLogMessage = "OrderPostFile Creation Exception -" + ex.Message + System.Environment.NewLine;
                                                            //    strExecutionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                            //    objCommon.WriteErrorLog(ex, strExecutionLogMessage);

                                                            //    ErrorResponse objErrorResponse = new ErrorResponse();
                                                            //    objErrorResponse.error = "Found exception while processing the file";
                                                            //    objErrorResponse.status = "Error";
                                                            //    objErrorResponse.code = "Excception while creating the order post file.";
                                                            //    string strErrorResponse = JsonConvert.SerializeObject(objErrorResponse);
                                                            //    DataSet dsFailureResponse = objCommon.jsonToDataSet(strErrorResponse);
                                                            //    dsFailureResponse.Tables[0].TableName = "Order-Create-Input";
                                                            //    objCommon.WriteDataToCsvFile(dsFailureResponse.Tables[0], fileName, strDatetime);
                                                            //    continue;
                                                            //}

                                                            //dtableOrderTemplate.TableName = "Template";

                                                            //   DataTable dtableOrderTemplate1 = new DataTable();
                                                            DataTable dtableOrderTemplateFinal = new DataTable();
                                                            // to filter the data based on the customer reference and finding the number of pieces 
                                                            if (customerName == objCommon.GetConfigValue("BBBCustomerName") && objCommon.GetConfigValue("BBBEnable_cartoncountsummary") == "Y")
                                                            {
                                                                int BBBEnableCartoncountsummaryOnly_orderlineitemAbove = int.Parse(objCommon.GetConfigValue("BBBEnableCartoncountsummaryOnly_orderlineitemAbove"));
                                                                if (dtableOrderTemplate.Rows.Count < BBBEnableCartoncountsummaryOnly_orderlineitemAbove)
                                                                {
                                                                    dtableOrderTemplateFinal = dtableOrderTemplate.Copy();
                                                                }
                                                                else
                                                                {
                                                                    DataView view = new DataView(dtableOrderTemplate);
                                                                    DataTable dtdistinctValues = view.ToTable(true, "Customer Reference");
                                                                    foreach (DataRow dr in dtdistinctValues.Rows)
                                                                    {
                                                                        object value = dr["Customer Reference"];
                                                                        if (value == DBNull.Value)
                                                                            break;
                                                                        string ReferenceId = Convert.ToString(dr["Customer Reference"]);
                                                                        try
                                                                        {
                                                                            if (dtableOrderTemplateFinal.Rows.Count > 0)
                                                                            {
                                                                                DataTable drresult = dtableOrderTemplate.Select("[Customer Reference]= '" + dr["Customer Reference"] + "'").CopyToDataTable();
                                                                                for (int row = 0; row < drresult.Rows.Count;)
                                                                                {
                                                                                    DataRow dr1 = dtableOrderTemplateFinal.NewRow();
                                                                                    for (int column = 0; column < drresult.Columns.Count; column++)
                                                                                    {
                                                                                        dr1[column] = drresult.Rows[row][column];
                                                                                    }
                                                                                    dr1["Pieces"] = drresult.Rows.Count;
                                                                                    dtableOrderTemplateFinal.Rows.Add(dr1.ItemArray);
                                                                                    break;
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                DataRow[] drresult = dtableOrderTemplate.Select("[Customer Reference]= '" + dr["Customer Reference"] + "'");
                                                                                var firstRow = drresult.AsEnumerable().First();
                                                                                //  firstRow.Table.Columns["Pieces"].DefaultValue = drresult.Length;
                                                                                //firstRow.ItemArray[14] = drresult.Length;
                                                                                // firstRow.AcceptChanges();
                                                                                dtableOrderTemplateFinal = new[] { firstRow }.CopyToDataTable();
                                                                                dtableOrderTemplateFinal.Rows[0]["Pieces"] = drresult.Length;
                                                                                dtableOrderTemplateFinal.AcceptChanges();
                                                                            }
                                                                        }
                                                                        catch (Exception ex)
                                                                        {
                                                                            executionLogMessage = "BBB summary file Creation Exception -" + ex.Message + System.Environment.NewLine;
                                                                            executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                                            objCommon.WriteExecutionLog(executionLogMessage);
                                                                            objCommon.WriteErrorLog(ex, executionLogMessage);

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                dtableOrderTemplateFinal = dtableOrderTemplate.Copy();
                                                            }

                                                            //IsSeen = true;
                                                            // To implement the logic for calculating the Billing and payable rates 

                                                            clsCommon.DSResponse objBPRatesResponse = new clsCommon.DSResponse();
                                                            objBPRatesResponse = objCommon.GetBillingRatesAndPayableRates_CustomerMappingDetails(company_no, customerNumber);
                                                            if (objBPRatesResponse.dsResp.ResponseVal)
                                                            {
                                                                DataTable dtBillingRates = new DataTable();
                                                                if (objBPRatesResponse.DS.Tables.Count > 0)
                                                                {
                                                                    dtBillingRates = objBPRatesResponse.DS.Tables[0].Copy();
                                                                }

                                                                DataTable dtPayableRates = new DataTable();
                                                                if (objBPRatesResponse.DS.Tables.Count > 1)
                                                                {
                                                                    dtPayableRates = objBPRatesResponse.DS.Tables[1].Copy();
                                                                }

                                                                // For CDS client Billing and payable amount calculation
                                                                if (customerName == objCommon.GetConfigValue("CDSCustomerName"))
                                                                {
                                                                    DataTable dtFSCRates = new DataTable();
                                                                    DataTable dtFSCRatesfromDB = new DataTable();
                                                                    DataTable tblFSCRatesFiltered = new DataTable();
                                                                    string strFscRateDetailsfilepath = objCommon.GetConfigValue("FSCRatesCustomerMappingFilepath");
                                                                    DataSet dsFscData = clsExcelHelper.ImportExcelXLSXToDataSet_FSCRATES(strFscRateDetailsfilepath, true, Convert.ToInt32(company_no), Convert.ToInt32(customerNumber));
                                                                    if (dsFscData != null && dsFscData.Tables[0].Rows.Count > 0)
                                                                    {
                                                                        dtFSCRates = dsFscData.Tables["FSCRatesMapping$"];
                                                                    }
                                                                    else
                                                                    {
                                                                        executionLogMessage = "Diesel price data not found in this file " + strFscRateDetailsfilepath + System.Environment.NewLine;
                                                                        executionLogMessage += "So not able to process this file, please update the fsc sheet with appropriate values";
                                                                        executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                                        objCommon.WriteExecutionLog(executionLogMessage);

                                                                        string fromMail = objCommon.GetConfigValue("FromMailID");
                                                                        string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                                                        string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                                                        string toMail = objCommon.GetConfigValue("CDSSendExceptionEmail");
                                                                        string subject = "Diesel price data not found in this file " + strFscRateDetailsfilepath;
                                                                        objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                                                                        throw new NullReferenceException("Diesel price data not found in this file " + strFscRateDetailsfilepath);
                                                                    }

                                                                    //DataTable dtCarrierFSCBillableRates = new DataTable();
                                                                    DataTable dtCarrierFSCRatesfromDB = new DataTable();
                                                                    DataTable tblCarrierFSCRatesFiltered = new DataTable();

                                                                    //string strCarrierFscRateDetailsfilepath = objCommon.GetConfigValue("CarrierFSCRatesCustomerMappingFilepath");
                                                                    //DataSet dsCarrierFscData = clsExcelHelper.ImportExcelXLSXToDataSet_FSCRATES(strCarrierFscRateDetailsfilepath, true, Convert.ToInt32(company_no), Convert.ToInt32(customerNumber));
                                                                    //if (dsCarrierFscData != null && dsCarrierFscData.Tables[0].Rows.Count > 0)
                                                                    //{
                                                                    //    dtCarrierFSCBillableRates = dsCarrierFscData.Tables["CarrierFSCRatesMapping$"];
                                                                    //}

                                                                    clsCommon.DSResponse objfscRatesResponse = new clsCommon.DSResponse();
                                                                    objfscRatesResponse = objCommon.GetFSCRates_MappingDetails(Convert.ToInt32(company_no), Convert.ToInt32(customerNumber));
                                                                    if (objfscRatesResponse.dsResp.ResponseVal)
                                                                    {
                                                                        if (objfscRatesResponse.DS.Tables.Count > 0)
                                                                        {
                                                                            dtFSCRatesfromDB = objfscRatesResponse.DS.Tables[0];
                                                                        }
                                                                        if (objfscRatesResponse.DS.Tables.Count > 1)
                                                                        {
                                                                            dtCarrierFSCRatesfromDB = objfscRatesResponse.DS.Tables[1];
                                                                        }
                                                                    }


                                                                    foreach (DataRow dr in dtableOrderTemplateFinal.Rows)
                                                                    {
                                                                        object value = dr["Delivery Date"];
                                                                        if (value == DBNull.Value)
                                                                            break;

                                                                        DateTime dtdeliveryDate = Convert.ToDateTime(Regex.Replace(value.ToString(), @"\t", ""));

                                                                        var invCulture = System.Globalization.CultureInfo.InvariantCulture;
                                                                        string deliveryName = Convert.ToString(dr["Customer Name"]);
                                                                        deliveryName = deliveryName.Replace("'", "");
                                                                        DataTable tblBillRatesFiltered = new DataTable();

                                                                        IEnumerable<DataRow> billratesfilteredRows = dtBillingRates.AsEnumerable()
                                                                        .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                                                                                 && row.Field<string>("DeliveryName") == deliveryName);

                                                                        if (billratesfilteredRows.Any())
                                                                        {
                                                                            tblBillRatesFiltered = billratesfilteredRows.CopyToDataTable();
                                                                        }
                                                                        else
                                                                        {
                                                                            billratesfilteredRows = dtBillingRates.AsEnumerable()
                                                                     .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                                                                     && row.Field<string>("DeliveryName") is null);

                                                                            if (billratesfilteredRows.Any())
                                                                            {
                                                                                tblBillRatesFiltered = billratesfilteredRows.CopyToDataTable();
                                                                            }
                                                                        }

                                                                        DataTable tblPayableRatesFiltered = new DataTable();
                                                                        IEnumerable<DataRow> payableratesfilteredRows = dtPayableRates.AsEnumerable()
                                                                        .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                                                                        && row.Field<string>("DeliveryName") == deliveryName);


                                                                        if (payableratesfilteredRows.Any())
                                                                        {
                                                                            tblPayableRatesFiltered = payableratesfilteredRows.CopyToDataTable();
                                                                        }
                                                                        else
                                                                        {
                                                                            payableratesfilteredRows = dtPayableRates.AsEnumerable()
                                                                        .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                                                                        && row.Field<string>("DeliveryName") is null);

                                                                            if (payableratesfilteredRows.Any())
                                                                            {
                                                                                tblPayableRatesFiltered = payableratesfilteredRows.CopyToDataTable();
                                                                            }
                                                                        }

                                                                        DataTable tblFSCBillRatesFiltered = new DataTable();
                                                                        double fscratePercentage = 0;
                                                                        double carrierfscratePercentage = 0;

                                                                        string fscratetype = string.Empty;
                                                                        string carrierfscratetype = string.Empty;

                                                                        IEnumerable<DataRow> fscbillratesfilteredRows = dtFSCRates.AsEnumerable()
                                                                        .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate)
                                                                        && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate")));

                                                                        if (fscbillratesfilteredRows.Any())
                                                                        {
                                                                            tblFSCBillRatesFiltered = fscbillratesfilteredRows.CopyToDataTable();

                                                                            decimal fuelcharge = 0;
                                                                            if (!string.IsNullOrEmpty(Convert.ToString(tblFSCBillRatesFiltered.Rows[0]["fuelcharge"])))
                                                                            {
                                                                                fuelcharge = Convert.ToDecimal(Convert.ToString(tblFSCBillRatesFiltered.Rows[0]["fuelcharge"]));
                                                                            }
                                                                            else
                                                                            {
                                                                                executionLogMessage = "Diesel price is missing for date  " + dtdeliveryDate.ToShortDateString() + System.Environment.NewLine;
                                                                                executionLogMessage += "So not able to process this file, please update the fsc sheet with appropriate values." + System.Environment.NewLine;
                                                                                executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                                                objCommon.WriteExecutionLog(executionLogMessage);
                                                                                string fromMail = objCommon.GetConfigValue("FromMailID");
                                                                                string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                                                                string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                                                                string toMail = objCommon.GetConfigValue("CDSSendExceptionEmail");
                                                                                string subject = "Diesel price is missing for date  " + dtdeliveryDate.ToShortDateString();
                                                                                objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                                                                                throw new NullReferenceException("Diesel price is missing for date  " + dtdeliveryDate.ToShortDateString());
                                                                            }

                                                                            if (dtFSCRatesfromDB.Rows.Count > 0)
                                                                            {
                                                                                IEnumerable<DataRow> fscratesfilteredRows = dtFSCRatesfromDB.AsEnumerable()
                                                                                .Where(row => (row.Field<decimal>("Start") <= fuelcharge) && (fuelcharge <= row.Field<decimal>("Stop"))
                                                                                 && row.Field<string>("DeliveryName") == deliveryName);

                                                                                if (fscratesfilteredRows.Any())
                                                                                {
                                                                                    tblFSCRatesFiltered = fscratesfilteredRows.CopyToDataTable();
                                                                                    fscratePercentage = Convert.ToDouble(tblFSCRatesFiltered.Rows[0]["Percent_FSCAMount"]);
                                                                                    fscratetype = Convert.ToString(tblFSCRatesFiltered.Rows[0]["Type"]);
                                                                                }
                                                                                else
                                                                                {

                                                                                    fscratesfilteredRows = dtFSCRatesfromDB.AsEnumerable()
                                                                                   .Where(row => (row.Field<decimal>("Start") <= fuelcharge) && (fuelcharge <= row.Field<decimal>("Stop"))
                                                                                    && row.Field<string>("DeliveryName") is null);
                                                                                    if (fscratesfilteredRows.Any())
                                                                                    {
                                                                                        tblFSCRatesFiltered = fscratesfilteredRows.CopyToDataTable();
                                                                                        fscratePercentage = Convert.ToDouble(tblFSCRatesFiltered.Rows[0]["Percent_FSCAMount"]);
                                                                                        fscratetype = Convert.ToString(tblFSCRatesFiltered.Rows[0]["Type"]);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        executionLogMessage = "FSC Billing Rates missing for this fuel charge   " + fuelcharge + System.Environment.NewLine;
                                                                                        executionLogMessage += "So not able to process this file, please update the billable rates mapping table with appropriate values";
                                                                                        executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                                                        objCommon.WriteExecutionLog(executionLogMessage);
                                                                                        string fromMail = objCommon.GetConfigValue("FromMailID");
                                                                                        string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                                                                        string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                                                                        string toMail = objCommon.GetConfigValue("ToMailID");
                                                                                        string subject = "FSC Billing Rates missing for this fuel charge   " + fuelcharge;
                                                                                        objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                                                                                        throw new NullReferenceException("Diesel prices missing for date  " + dtdeliveryDate);
                                                                                    }
                                                                                }
                                                                            }
                                                                            if (dtCarrierFSCRatesfromDB.Rows.Count > 0)
                                                                            {
                                                                                IEnumerable<DataRow> CarrierfscratesfilteredRows = dtCarrierFSCRatesfromDB.AsEnumerable()
                                                                                .Where(row => (row.Field<decimal>("Start") <= fuelcharge)
                                                                                && (fuelcharge <= row.Field<decimal>("Stop"))
                                                                                && row.Field<string>("DeliveryName") == deliveryName);

                                                                                if (CarrierfscratesfilteredRows.Any())
                                                                                {
                                                                                    tblCarrierFSCRatesFiltered = CarrierfscratesfilteredRows.CopyToDataTable();
                                                                                    carrierfscratePercentage = Convert.ToDouble(tblCarrierFSCRatesFiltered.Rows[0]["Percent_FSCAMount"]);
                                                                                    carrierfscratetype = Convert.ToString(tblCarrierFSCRatesFiltered.Rows[0]["Type"]);
                                                                                }
                                                                                else
                                                                                {
                                                                                    CarrierfscratesfilteredRows = dtCarrierFSCRatesfromDB.AsEnumerable()
                                                                                .Where(row => (row.Field<decimal>("Start") <= fuelcharge)
                                                                                && (fuelcharge <= row.Field<decimal>("Stop"))
                                                                                && row.Field<string>("DeliveryName") is null);
                                                                                    if (CarrierfscratesfilteredRows.Any())
                                                                                    {
                                                                                        tblCarrierFSCRatesFiltered = CarrierfscratesfilteredRows.CopyToDataTable();
                                                                                        carrierfscratePercentage = Convert.ToDouble(tblCarrierFSCRatesFiltered.Rows[0]["Percent_FSCAMount"]);
                                                                                        carrierfscratetype = Convert.ToString(tblCarrierFSCRatesFiltered.Rows[0]["Type"]);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        executionLogMessage = "FSC Payable Rates missing for this fuel charge   " + fuelcharge + System.Environment.NewLine;
                                                                                        executionLogMessage += "So not able to process this file, please update the payable rates mapping table with appropriate values";
                                                                                        executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                                                        objCommon.WriteExecutionLog(executionLogMessage);
                                                                                        string fromMail = objCommon.GetConfigValue("FromMailID");
                                                                                        string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                                                                        string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                                                                        string toMail = objCommon.GetConfigValue("ToMailID");
                                                                                        string subject = "FSC Billing Rates missing for this fuel charge   " + fuelcharge;
                                                                                        objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                                                                                        throw new NullReferenceException("Diesel prices missing for date  " + dtdeliveryDate);
                                                                                    }
                                                                                }
                                                                            }
                                                                        }

                                                                        if (tblBillRatesFiltered.Rows.Count > 0)
                                                                        {
                                                                            double billRate = 0;
                                                                            double minimunRate = 0; ;

                                                                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["minimun_rate"])))
                                                                            {
                                                                                minimunRate = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["minimun_rate"]); ;
                                                                            }

                                                                            if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                                                                            {
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"])))
                                                                                {
                                                                                    billRate = Convert.ToDouble(1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]));
                                                                                    if (billRate < minimunRate)
                                                                                    {
                                                                                        billRate = minimunRate;
                                                                                    }
                                                                                    dr["Bill Rate"] = Math.Round(Convert.ToDouble(billRate), 2);
                                                                                }

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"])))
                                                                                    dr["rate_buck_amt2"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"])))
                                                                                    dr["Pieces ACC"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(fscratePercentage)))
                                                                                {
                                                                                    if (fscratetype.ToString().ToUpper() == "F")
                                                                                    {
                                                                                        dr["FSC"] = Math.Round(Convert.ToDouble(fscratePercentage), 2);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        dr["FSC"] = Math.Round(Convert.ToDouble(billRate * fscratePercentage) / 100, 2);
                                                                                    }
                                                                                }

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"])))
                                                                                    dr["rate_buck_amt4"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"])))
                                                                                    dr["rate_buck_amt5"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"])))
                                                                                    dr["rate_buck_amt6"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"])))
                                                                                    dr["rate_buck_amt7"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"])))
                                                                                    dr["rate_buck_amt8"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"])))
                                                                                    dr["rate_buck_amt9"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"])))
                                                                                    dr["rate_buck_amt11"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"]);

                                                                            }
                                                                            else
                                                                            {
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"])))
                                                                                {
                                                                                    billRate = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]);
                                                                                    if (billRate < minimunRate)
                                                                                    {
                                                                                        billRate = minimunRate;
                                                                                    }
                                                                                    dr["Bill Rate"] = Math.Round(Convert.ToDouble(billRate), 2);
                                                                                }

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"])))
                                                                                    dr["rate_buck_amt2"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"])))
                                                                                    dr["Pieces ACC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(fscratePercentage)))
                                                                                {
                                                                                    //  dr["FSC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"]);

                                                                                    if (fscratetype.ToString().ToUpper() == "F")
                                                                                    {
                                                                                        dr["FSC"] = Math.Round(Convert.ToDouble(fscratePercentage), 2);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        dr["FSC"] = Math.Round(Convert.ToDouble(billRate * fscratePercentage) / 100, 2);
                                                                                    }
                                                                                }

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"])))
                                                                                    dr["rate_buck_amt4"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"])))
                                                                                    dr["rate_buck_amt5"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"])))
                                                                                    dr["rate_buck_amt6"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"])))
                                                                                    dr["rate_buck_amt7"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"])))
                                                                                    dr["rate_buck_amt8"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"])))
                                                                                    dr["rate_buck_amt9"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"])))
                                                                                    dr["rate_buck_amt11"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"]);

                                                                            }
                                                                        }

                                                                        if (tblPayableRatesFiltered.Rows.Count > 0)
                                                                        {
                                                                            double carrierBasePay = 0;
                                                                            double minimunRate = 0;
                                                                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["minimun_rate"])))
                                                                            {
                                                                                minimunRate = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["minimun_rate"]);
                                                                            }

                                                                            if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                                                                            {
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                                                                                {
                                                                                    carrierBasePay = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);
                                                                                    if (carrierBasePay < minimunRate)
                                                                                    {
                                                                                        carrierBasePay = minimunRate;
                                                                                    }
                                                                                    dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);
                                                                                }

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge5"])))
                                                                                    dr["Carrier ACC"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge5"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(carrierfscratePercentage)))
                                                                                {
                                                                                    //  dr["Carrier FSC"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge6"]);
                                                                                    if (carrierfscratetype.ToString().ToUpper() == "F")
                                                                                    {
                                                                                        dr["Carrier FSC"] = Math.Round(Convert.ToDouble(carrierfscratePercentage), 2);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        dr["Carrier FSC"] = Math.Round(Convert.ToDouble(carrierBasePay * carrierfscratePercentage) / 100, 2);
                                                                                    }
                                                                                }

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge2"])))
                                                                                    dr["charge2"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge2"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge3"])))
                                                                                    dr["charge3"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge3"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge4"])))
                                                                                    dr["charge4"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge4"]);

                                                                            }
                                                                            else
                                                                            {
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                                                                                {
                                                                                    // dr["Carrier Base Pay"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);
                                                                                    carrierBasePay = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);
                                                                                    if (carrierBasePay < minimunRate)
                                                                                    {
                                                                                        carrierBasePay = minimunRate;
                                                                                    }
                                                                                    dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);
                                                                                }
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge5"])))
                                                                                    dr["Carrier ACC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge5"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(carrierfscratePercentage)))
                                                                                {
                                                                                    //dr["Carrier FSC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge6"]);

                                                                                    if (carrierfscratetype.ToString().ToUpper() == "F")
                                                                                    {
                                                                                        dr["Carrier FSC"] = Math.Round(Convert.ToDouble(carrierfscratePercentage), 2);
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        dr["Carrier FSC"] = Math.Round(Convert.ToDouble(carrierBasePay * carrierfscratePercentage) / 100, 2);
                                                                                    }
                                                                                }
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge2"])))
                                                                                    dr["charge2"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge2"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge3"])))
                                                                                    dr["charge3"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge3"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge4"])))
                                                                                    dr["charge4"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge4"]);

                                                                            }
                                                                        }

                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    foreach (DataRow dr in dtableOrderTemplateFinal.Rows)
                                                                    {
                                                                        object value = dr["Delivery Date"];
                                                                        if (value == DBNull.Value)
                                                                            break;

                                                                        DateTime dtdeliveryDate = Convert.ToDateTime(Regex.Replace(value.ToString(), @"\t", ""));

                                                                        var invCulture = System.Globalization.CultureInfo.InvariantCulture;

                                                                        DataTable tblBillRatesFiltered = new DataTable();
                                                                        IEnumerable<DataRow> billratesfilteredRows = dtBillingRates.AsEnumerable()
                                                                        .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate")));

                                                                        if (billratesfilteredRows.Any())
                                                                        {
                                                                            tblBillRatesFiltered = billratesfilteredRows.CopyToDataTable();
                                                                        }

                                                                        DataTable tblPayableRatesFiltered = new DataTable();
                                                                        IEnumerable<DataRow> payableratesfilteredRows = dtPayableRates.AsEnumerable()
                                                                        .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate")));


                                                                        if (payableratesfilteredRows.Any())
                                                                        {
                                                                            tblPayableRatesFiltered = payableratesfilteredRows.CopyToDataTable();
                                                                        }

                                                                        if (tblBillRatesFiltered.Rows.Count > 0)
                                                                        {
                                                                            if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                                                                            {
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"])))
                                                                                    dr["Bill Rate"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"])))
                                                                                    dr["rate_buck_amt2"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"])))
                                                                                    dr["Pieces ACC"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"])))
                                                                                    dr["FSC"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"]);


                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"])))
                                                                                    dr["rate_buck_amt4"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"])))
                                                                                    dr["rate_buck_amt5"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"])))
                                                                                    dr["rate_buck_amt6"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"])))
                                                                                    dr["rate_buck_amt7"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"])))
                                                                                    dr["rate_buck_amt8"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"])))
                                                                                    dr["rate_buck_amt9"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"])))
                                                                                    dr["rate_buck_amt11"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"]);

                                                                            }
                                                                            else
                                                                            {
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"])))
                                                                                    dr["Bill Rate"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"])))
                                                                                    dr["rate_buck_amt2"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"])))
                                                                                    dr["Pieces ACC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"])))
                                                                                    dr["FSC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"]);


                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"])))
                                                                                    dr["rate_buck_amt4"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"])))
                                                                                    dr["rate_buck_amt5"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"])))
                                                                                    dr["rate_buck_amt6"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"])))
                                                                                    dr["rate_buck_amt7"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"])))
                                                                                    dr["rate_buck_amt8"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"])))
                                                                                    dr["rate_buck_amt9"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"])))
                                                                                    dr["rate_buck_amt11"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"]);

                                                                            }
                                                                        }

                                                                        if (tblPayableRatesFiltered.Rows.Count > 0)
                                                                        {
                                                                            if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                                                                            {
                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                                                                                    dr["Carrier Base Pay"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge5"])))
                                                                                    dr["Carrier ACC"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge5"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge6"])))
                                                                                    dr["Carrier FSC"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge6"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge2"])))
                                                                                    dr["charge2"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge2"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge3"])))
                                                                                    dr["charge3"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge3"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge4"])))
                                                                                    dr["charge4"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge4"]);

                                                                            }
                                                                            else
                                                                            {

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                                                                                    dr["Carrier Base Pay"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge5"])))
                                                                                    dr["Carrier ACC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge5"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge6"])))
                                                                                    dr["Carrier FSC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge6"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge2"])))
                                                                                    dr["charge2"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge2"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge3"])))
                                                                                    dr["charge3"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge3"]);

                                                                                if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge4"])))
                                                                                    dr["charge4"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge4"]);

                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            IsSeen = true;
                                                            if (!info.Read && IsSeen)
                                                            {
                                                                oClient.MarkAsRead(info, true);
                                                                executionLogMessage = "Mark email as read, From :  " + oMail.From.ToString() + " , ReceivedDate" + oMail.ReceivedDate;
                                                                objCommon.WriteExecutionLog(executionLogMessage);
                                                            }

                                                            dtableOrderTemplateFinal.TableName = "Template";
                                                            clsExcelHelper.ExportDataToXLSX(dtableOrderTemplateFinal, attachmentPath, fileName);


                                                            if (customerName == objCommon.GetConfigValue("BBBCustomerName") && productSubCode == "TND")
                                                            {
                                                                DataTable dtConfiguredData = new DataTable();
                                                                objDsRes = new clsCommon.DSResponse();
                                                                productSubCode = "DEL";
                                                                objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);
                                                                if (objDsRes.dsResp.ResponseVal)
                                                                {
                                                                    dtConfiguredData = objDsRes.DS.Tables[0];
                                                                    customerNumber = Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);

                                                                    dsExcel = ConvertBBBDELFlatfileToDataTable(tempfilePath + @"\" + fileName,
                                                                        dtConfiguredData);
                                                                    dtableOrderTemplate.Clear();
                                                                    dtOrderData.Clear();
                                                                    dtOrderData = dsExcel.Tables[0];
                                                                    GenerateOrderTemplate(ref dtableOrderTemplate, dtOrderData, fileName, customerName, locationCode, productCode, productSubCode);
                                                                    ProcessBBBDELfileToDataTable(fileName, dtableOrderTemplate, Convert.ToInt32(company_no), Convert.ToInt32(customerNumber), attachmentPath);
                                                                }
                                                            }
                                                            objCommon.CleanAttachmentWorkingFolder();

                                                        }
                                                        else
                                                        {
                                                            executionLogMessage = "Order Post Template Mapping Missing " + System.Environment.NewLine;
                                                            executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                                            executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                                            executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                                            executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                                            emailSubject = "Order Post Template Mapping Missing";
                                                            objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                                            objCommon.WriteExecutionLog(executionLogMessage);
                                                        }

                                                        break;
                                                    default:
                                                        executionLogMessage = "This product Excel Template Mapping Not Implemented, Please find the details below and implement the same " + System.Environment.NewLine;
                                                        executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                                        executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                                        executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                                        executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                                        emailSubject = "This product Excel Template Mapping Not Implemented";
                                                        objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                                        objCommon.WriteExecutionLog(executionLogMessage);
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                executionLogMessage = "No data found after Export, Please check the file " + System.Environment.NewLine;
                                                executionLogMessage += "Attachment Name -" + fileName + System.Environment.NewLine;
                                                executionLogMessage = "From :  " + oMail.From.ToString() + System.Environment.NewLine;
                                                executionLogMessage = "Email Address :  " + oMail.From.Address.ToString() + System.Environment.NewLine;
                                                executionLogMessage += "Subject : " + oMail.Subject + System.Environment.NewLine;
                                                executionLogMessage += "ReceivedDate : " + oMail.ReceivedDate + System.Environment.NewLine;
                                                emailSubject = "No data found for this Excel Post Convertion to data set";
                                                objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                                objCommon.WriteExecutionLog(executionLogMessage);
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            executionLogMessage = "Email Mapping Missing For this Customer" + System.Environment.NewLine;
                                            executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                            executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                            executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                            executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                            emailSubject = "Email Mapping Missing For " + customerName;
                                            objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                            objCommon.WriteExecutionLog(executionLogMessage);
                                            continue;
                                        }

                                        // mark unread email as read, next time this email won't be retrieved again
                                        //if (!info.Read && IsSeen)
                                        //{
                                        //    oClient.MarkAsRead(info, true);
                                        //    //  Console.WriteLine("Mark email as read\r\n");
                                        //    strExecutionLogMessage = "Mark email as read, From :  " + oMail.From.ToString() + " , ReceivedDate" + oMail.ReceivedDate;
                                        //    objCommon.WriteExecutionLog(strExecutionLogMessage);
                                        //}
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Console.WriteLine(ex.Message);
                                    executionLogMessage = "Found Exception - Exception while processing the attachment for the below mentioned email" + System.Environment.NewLine;
                                    executionLogMessage += "Email From -" + oMail.From.ToString() + System.Environment.NewLine;
                                    executionLogMessage += "Email Address -" + oMail.From.Address.ToString() + System.Environment.NewLine;
                                    executionLogMessage += "Email Subject -" + oMail.Subject.ToString() + System.Environment.NewLine;
                                    executionLogMessage += "Email ReceivedDate -" + oMail.ReceivedDate + System.Environment.NewLine;
                                    if (ex.Message.Contains("Index was outside the bounds of the array"))
                                    {
                                        executionLogMessage += "Exception There may be issue while reading invalid formated attchement/file -" + fileName + System.Environment.NewLine;
                                    }
                                    objCommon.WriteErrorLog(ex, "StartProcessing - Exception while processing the attachment/file - " + fileName, executionLogMessage);
                                }
                            }

                            //if (!info.Read)
                            //{
                            //    oClient.MarkAsRead(info, true);
                            //}
                        }
                        //}


                    }


                    // Generate an unqiue email file name based on date time.
                    //  string fileName = _generateFileName(i + 1);
                    //  string fullPath = string.Format("{0}\\{1}", localInbox, fileName);

                    // Save email to local disk
                    //oMail.SaveAs(fullPath, true);



                    // Mark email as deleted from IMAP4 server.
                    // oClient.Delete(info);
                }

                // Quit and expunge emails marked as deleted from IMAP4 server.
                try
                {
                    oClient.Quit();
                }
                catch (Exception ex)
                {

                }
                // Console.WriteLine("Completed!");
                executionLogMessage = "Process Completed";
                objCommon.WriteExecutionLog(executionLogMessage);
            }
            catch (Exception ex)
            {
                // Console.WriteLine(ex.Message);
                objCommon.WriteErrorLog(ex, "StartProcessing - Exception Occurred while Processing");
            }
        }

        private static string Right(this string str, int n)
        {
            if (n > str.Length)
            {
                return str;
            }

            return str.Substring(str.Length - n);
        }
        private static DataSet ConvertBBBINBFlatfileToDataTable(string filePath, DataTable dtConfiguredData)
        {
            //DataTable tbl = new DataTable();
            DataSet output = new DataSet();
            DataTable tbl = new DataTable();
            tbl.Clear();

            tbl.Columns.Add("TRLR_NUM");
            tbl.Columns.Add("CLOSE_DT");
            tbl.Columns.Add("ARV_DT");
            tbl.Columns.Add("CARRIER");
            tbl.Columns.Add("BOL");
            tbl.Columns.Add("CAR_PRO");
            tbl.Columns.Add("CART");
            tbl.Columns.Add("WEIGHT");
            tbl.Columns.Add("CUBE");
            tbl.Columns.Add("CLASS");
            tbl.Columns.Add("ORIG_SHIP_ID");
            tbl.Columns.Add("SHIP_TYPE");
            tbl.Columns.Add("SHIP_ID");
            tbl.Columns.Add("SHIP_NAME");
            tbl.Columns.Add("SHIP_ADDR_1");
            tbl.Columns.Add("SHIP_ADDR_2");
            tbl.Columns.Add("SHIP_ADDR_3");
            tbl.Columns.Add("SHIP_CITY");
            tbl.Columns.Add("SHIP_STATE");
            tbl.Columns.Add("SHIP_ZIP");
            tbl.Columns.Add("Billing_Customer_Number");
            tbl.Columns.Add("Service_Type");
            tbl.Columns.Add("Entered_by");
            tbl.Columns.Add("Pickup_Delivery_Transfer_Flag");
            tbl.Columns.Add("dim_height");
            tbl.Columns.Add("dim_length");
            tbl.Columns.Add("dim_width");
            tbl.Columns.Add("pickup_name");
            tbl.Columns.Add("pickup_attention");
            tbl.Columns.Add("deliver_attention");
            tbl.Columns.Add("delivery_address");
            tbl.Columns.Add("delivery_city");
            tbl.Columns.Add("delivery_state");
            tbl.Columns.Add("delivery_zip");

            tbl.Columns.Add("AddressValue");
            tbl.Columns.Add("CityValue");
            tbl.Columns.Add("StateValue");
            tbl.Columns.Add("ZipValue");

            tbl.Columns.Add("pickup_name_Value");
            tbl.Columns.Add("Pickup_will_be_ready_by_Value");
            tbl.Columns.Add("Pickup_no_later_than_Value");
            tbl.Columns.Add("Pickup_actual_arrival_time_Value");
            tbl.Columns.Add("Pickup_actual_depart_time_Value");
            tbl.Columns.Add("Deliver_no_earlier_than_Value");
            tbl.Columns.Add("Deliver_no_later_than_Value");
            tbl.Columns.Add("Delivery_actual_arrive_time_Value");
            tbl.Columns.Add("Delivery_actual_depart_time_Value");
            tbl.Columns.Add("CustomerName_Value");
            tbl.Columns.Add("Correct_Driver_Number_Value");
            tbl.Columns.Add("Delivery_text_signature_Value");

            //for (int col = 0; col < numberOfColumns; col++)
            //    tbl.Columns.Add(new DataColumn("Column" + (col + 1).ToString()));

            string[] lines = System.IO.File.ReadAllLines(filePath);
            lines = lines.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            foreach (string line in lines)
            {
                //Console.WriteLine("line length" + line.Length);

                DataRow _newRow = tbl.NewRow();
                _newRow["TRLR_NUM"] = "\t" + line.Substring(0, 12).Trim();
                //_newRow["CLOSE_DT"] =  line.Substring(12, 8);
                _newRow["CLOSE_DT"] = "\t" + DateTime.ParseExact(line.Substring(12, 8), "yyyyMMdd",
                            CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");

                _newRow["ARV_DT"] = "\t" + DateTime.ParseExact(line.Substring(20, 8), "yyyyMMdd",
                           CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");

                _newRow["CARRIER"] = "\t" + line.Substring(28, 4).Trim();

                string strbol = line.Substring(32, 30).Trim();
                // _newRow["BOL"] = "\t" + line.Substring(32, 30).Trim();
                _newRow["BOL"] = "\t" + Right(strbol, 23);
                _newRow["SHIP_NAME"] = "\t" + strbol;

                string strCarPro = line.Substring(62, 30).Trim();
                // _newRow["CAR_PRO"] = "\t" + line.Substring(62, 30).Trim();
                _newRow["CAR_PRO"] = "\t" + Right(strCarPro, 15);
                _newRow["pickup_name"] = "\t" + strCarPro;


                _newRow["CART"] = "\t" + line.Substring(92, 28).Trim();
                _newRow["WEIGHT"] = "\t" + line.Substring(120, 12).Trim();
                _newRow["CUBE"] = "\t" + line.Substring(132, 12).Trim();
                _newRow["dim_height"] = "\t" + line.Substring(132, 12).Trim();
                _newRow["dim_length"] = "\t" + line.Substring(132, 12).Trim();
                _newRow["dim_width"] = "\t" + line.Substring(132, 12).Trim();
                _newRow["CLASS"] = "\t" + line.Substring(144, 7).Trim();
                _newRow["ORIG_SHIP_ID"] = "\t" + line.Substring(151, 11).Trim();
                _newRow["SHIP_TYPE"] = "\t" + line.Substring(162, 1).Trim();
                if (!string.IsNullOrEmpty(line.Substring(163, 8).Trim()))
                {
                    _newRow["SHIP_ID"] = "\t" + line.Substring(163, 8).Trim();
                }
                else
                {
                    _newRow["SHIP_ID"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["RequestedByValue"]);
                }
                //_newRow["SHIP_NAME"] = "\t" + line.Substring(171, 30).Trim();
                string strshipname = line.Substring(171, 30).Trim();
                string strpickup_attention = "";
                string strdeliver_attention = "";

                if (strshipname.Length > 15)
                {
                    strpickup_attention = "\t" + strshipname.Substring(0, 15);
                    strdeliver_attention = "\t" + strshipname.Substring(15, strshipname.Length - 15);
                }
                else
                {
                    strpickup_attention = "\t" + strshipname;
                    strdeliver_attention = "";
                }

                _newRow["pickup_attention"] = strpickup_attention;
                _newRow["deliver_attention"] = "\t" + strdeliver_attention;

                if (!string.IsNullOrEmpty(line.Substring(201, 30).Trim()))
                {
                    _newRow["SHIP_ADDR_1"] = "\t" + line.Substring(201, 30).Trim();
                }
                else
                {
                    _newRow["SHIP_ADDR_1"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                }
                _newRow["SHIP_ADDR_2"] = "\t" + line.Substring(231, 30).Trim();
                _newRow["SHIP_ADDR_3"] = "\t" + line.Substring(261, 30).Trim();
                if (!string.IsNullOrEmpty(line.Substring(291, 20).Trim()))
                {
                    _newRow["SHIP_CITY"] = "\t" + line.Substring(291, 20).Trim();
                }
                else
                {
                    _newRow["SHIP_CITY"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]);
                }
                if (!string.IsNullOrEmpty(line.Substring(311, 2).Trim()))
                {
                    _newRow["SHIP_STATE"] = "\t" + line.Substring(311, 2).Trim();
                }
                else
                {
                    _newRow["SHIP_STATE"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                }
                if (!string.IsNullOrEmpty(line.Substring(313, 10).Trim()))
                {
                    _newRow["SHIP_ZIP"] = "\t" + line.Substring(313, 10).Trim();
                }
                else
                {
                    _newRow["SHIP_ZIP"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);
                }
                _newRow["Billing_Customer_Number"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                _newRow["Service_Type"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ServiceTypeValue"]);
                _newRow["Entered_by"] = Convert.ToString(dtConfiguredData.Rows[0]["EntredByValue"]);
                _newRow["Pickup_Delivery_Transfer_Flag"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["PickupDeliveryTransferFlagValue"]);
                _newRow["delivery_address"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                _newRow["delivery_city"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]); ;
                _newRow["delivery_state"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                _newRow["delivery_zip"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);

                _newRow["AddressValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                _newRow["CityValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]);
                _newRow["StateValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                _newRow["ZipValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);

                _newRow["pickup_name_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["pickup_name_Value"]);
                _newRow["Pickup_will_be_ready_by_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_will_be_ready_by_Value"]);
                _newRow["Pickup_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_no_later_than_Value"]);
                _newRow["Pickup_actual_arrival_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_arrival_time_Value"]);
                _newRow["Pickup_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_depart_time_Value"]);
                _newRow["Deliver_no_earlier_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_earlier_than_Value"]);
                _newRow["Deliver_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_later_than_Value"]);
                _newRow["Delivery_actual_arrive_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_actual_arrive_time_Value"]);
                _newRow["Delivery_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_actual_depart_time_Value"]);
                _newRow["CustomerName_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["CustomerName_Value"]);
                _newRow["Correct_Driver_Number_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Correct_Driver_Number_Value"]);
                _newRow["Delivery_text_signature_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_text_signature_Value"]);


                tbl.Rows.Add(_newRow);
            }
            output.Tables.Add(tbl);
            return output;
        }

        private static DataTable ConvertBBBFlatfileToDataTable(string filePath, int numberOfColumns)
        {
            DataTable tbl = new DataTable();

            for (int col = 0; col < numberOfColumns; col++)
                tbl.Columns.Add(new DataColumn("Column" + (col + 1).ToString()));

            string[] lines = System.IO.File.ReadAllLines(filePath);
            lines = lines.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            foreach (string line in lines)
            {
                var cols = line.Split(' ');
                cols = cols.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                DataRow dr = tbl.NewRow();
                for (int cIndex = 0; cIndex < numberOfColumns; cIndex++)
                {
                    dr[cIndex] = cols[cIndex];
                }
                tbl.Rows.Add(dr);
            }
            return tbl;
        }

        private static DataSet TGTGenerateOrderDataTable_Databaselevel(DataTable dtinputData, DataTable dtConfiguredData, string type)
        {

            string strExecutionLogMessage = string.Empty;
            clsCommon objCommon = new clsCommon();
            DataSet output = new DataSet();
            DataTable dtableOrderTemplate = new DataTable();
            dtableOrderTemplate.Clear();

            dtableOrderTemplate.Columns.Add("Delivery_Date");
            dtableOrderTemplate.Columns.Add("Billing_Customer_Number");
            dtableOrderTemplate.Columns.Add("Customer_Reference");
            dtableOrderTemplate.Columns.Add("ServiceTypeValue");
            dtableOrderTemplate.Columns.Add("EnteredByValue");
            dtableOrderTemplate.Columns.Add("RequestedByValue");
            dtableOrderTemplate.Columns.Add("PickupDeliveryTransferFlagValue");

            dtableOrderTemplate.Columns.Add("pickup_name");
            dtableOrderTemplate.Columns.Add("delivery_name");

            foreach (DataRow dr in dtinputData.Rows)
            {
                object value = dr["Load ID"];
                if (value == DBNull.Value)
                {
                    dr.Delete();
                    break;
                }
                clsCommon.DSResponse objDsResponse = new clsCommon.DSResponse();

                string strCustomer_Reference = Convert.ToString(dr["Load ID"]);
                string strOriginAddress = Convert.ToString(dr["Origin Address"]).Replace("\"", string.Empty);
                string strDestinationAddress = Convert.ToString(dr["Destination Address"]).Replace("\"", string.Empty); ;
                objDsResponse = objCommon.GetTGTCustomerMappingDetails(type, strOriginAddress, strDestinationAddress);
                if (objDsResponse.dsResp.ResponseVal)
                {
                    if (objDsResponse.DS.Tables[0].Rows.Count > 0 && objDsResponse.DS.Tables[1].Rows.Count > 0)
                    {
                        DataRow _newRow = dtableOrderTemplate.NewRow();
                        DateTime Delivery_Date = Convert.ToDateTime(dr["Load End Date/Expected Delivery Date"]);
                        _newRow["Delivery_Date"] = Delivery_Date.ToString("MM/dd/yyyy");
                        _newRow["Billing_Customer_Number"] = Convert.ToString(objDsResponse.DS.Tables[0].Rows[0]["BillingCustomerNumber"]);
                        _newRow["Customer_Reference"] = Convert.ToString(dr["Load ID"]);
                        _newRow["ServiceTypeValue"] = Convert.ToString(dtConfiguredData.Rows[0]["ServiceTypeValue"]);
                        _newRow["EnteredByValue"] = Convert.ToString(dtConfiguredData.Rows[0]["EntredByValue"]);
                        _newRow["RequestedByValue"] = Convert.ToString(dtConfiguredData.Rows[0]["RequestedByValue"]);
                        _newRow["PickupDeliveryTransferFlagValue"] = Convert.ToString(dtConfiguredData.Rows[0]["PickupDeliveryTransferFlagValue"]);
                        _newRow["pickup_name"] = Convert.ToString(objDsResponse.DS.Tables[1].Rows[0]["PickName_DeliveyName"]);
                        _newRow["delivery_name"] = Convert.ToString(objDsResponse.DS.Tables[1].Rows[0]["PickName_DeliveyName"]);
                        dtableOrderTemplate.Rows.Add(_newRow);
                    }
                    else
                    {
                        strExecutionLogMessage = "TGT Customer NBR/Route NBR not found for this Customer Reference " + System.Environment.NewLine;
                        strExecutionLogMessage += "Customer Reference -" + strCustomer_Reference + System.Environment.NewLine;
                        strExecutionLogMessage += "Type -" + type + System.Environment.NewLine;
                        strExecutionLogMessage += "Origin Address -" + strOriginAddress + System.Environment.NewLine;
                        strExecutionLogMessage += "Destination Address -" + strDestinationAddress + System.Environment.NewLine;
                        objCommon.WriteExecutionLog(strExecutionLogMessage);
                    }
                }
            }
            output.Tables.Add(dtableOrderTemplate);
            return output;
        }

        private static DataSet TGTGenerateOrderDataTable(DataTable dtinputData, DataTable dtConfiguredData, string type)
        {


            string strExecutionLogMessage = string.Empty;
            clsCommon objCommon = new clsCommon();
            DataSet output = new DataSet();
            DataTable dtableOrderTemplate = new DataTable();
            dtableOrderTemplate.Clear();

            dtableOrderTemplate.Columns.Add("Delivery_Date");
            dtableOrderTemplate.Columns.Add("Billing_Customer_Number");
            dtableOrderTemplate.Columns.Add("Customer_Reference");
            dtableOrderTemplate.Columns.Add("ServiceTypeValue");
            dtableOrderTemplate.Columns.Add("EnteredByValue");
            dtableOrderTemplate.Columns.Add("RequestedByValue");
            dtableOrderTemplate.Columns.Add("PickupDeliveryTransferFlagValue");

            dtableOrderTemplate.Columns.Add("pickup_name");
            dtableOrderTemplate.Columns.Add("delivery_name");

            DataSet dsExcel = new DataSet();
            string TgtCustomerMappingFilepath = objCommon.GetConfigValue("TgtCustomerMappingFilepath");

            dsExcel = clsExcelHelper.ImportExcelXLSXToDataSet(TgtCustomerMappingFilepath, false);

            if (dsExcel != null && dsExcel.Tables[0].Rows.Count > 0)
            {
                DataSet dsResult = new DataSet();
                DataTable dtresultCustomerNbrMapping = new DataTable();
                DataTable dtresultRouteNbrMapping = new DataTable();

                DataTable dataTableCustomerNbrMapping = dsExcel.Tables["CustomerNbrMapping$"];
                DataTable dataTableRouteNbrMapping = dsExcel.Tables["RouteNbrMapping$"];

                foreach (DataRow dr in dtinputData.Rows)
                {
                    object value = dr["Load ID"];
                    if (value == DBNull.Value)
                    {
                        dr.Delete();
                        continue;
                    }

                    string strCustomer_Reference = Convert.ToString(dr["Load ID"]);
                    string strOriginAddress = Convert.ToString(dr["Origin Address"]).Replace("\"", string.Empty);
                    string strDestinationAddress = Convert.ToString(dr["Destination Address"]).Replace("\"", string.Empty); ;

                    var rowsSummary = dataTableCustomerNbrMapping.AsEnumerable().Where(x => x.Field<string>("Type") == type && x.Field<string>("Address") == strDestinationAddress);
                    if (rowsSummary.Any())
                    {
                        dtresultCustomerNbrMapping = rowsSummary.CopyToDataTable();
                    }
                    else
                    {
                        strExecutionLogMessage = "TGT Customer NBR not found for this Customer Reference " + System.Environment.NewLine;
                        strExecutionLogMessage += "Customer Reference -" + strCustomer_Reference + System.Environment.NewLine;
                        strExecutionLogMessage += "Type -" + type + System.Environment.NewLine;
                        strExecutionLogMessage += "Origin Address -" + strOriginAddress + System.Environment.NewLine;
                        strExecutionLogMessage += "Destination Address -" + strDestinationAddress + System.Environment.NewLine;
                        objCommon.WriteExecutionLog(strExecutionLogMessage);
                        objCommon.WriteExecutionLog(strExecutionLogMessage);
                        continue;
                    }

                    var rowsDetails = dataTableRouteNbrMapping.AsEnumerable().Where(x => x.Field<string>("Type") == type && x.Field<string>("Address") == strOriginAddress);
                    if (rowsDetails.Any())
                    {
                        dtresultRouteNbrMapping = rowsDetails.CopyToDataTable();
                    }
                    else
                    {
                        strExecutionLogMessage = "TGT Route NBR not found for this Customer Reference " + System.Environment.NewLine;
                        strExecutionLogMessage += "Customer Reference -" + strCustomer_Reference + System.Environment.NewLine;
                        strExecutionLogMessage += "Type -" + type + System.Environment.NewLine;
                        strExecutionLogMessage += "Origin Address -" + strOriginAddress + System.Environment.NewLine;
                        strExecutionLogMessage += "Destination Address -" + strDestinationAddress + System.Environment.NewLine;
                        objCommon.WriteExecutionLog(strExecutionLogMessage);
                        objCommon.WriteExecutionLog(strExecutionLogMessage);
                        continue;
                    }


                    if (dtresultCustomerNbrMapping.Rows.Count > 0 && dtresultRouteNbrMapping.Rows.Count > 0)
                    {
                        DataRow _newRow = dtableOrderTemplate.NewRow();
                        DateTime Delivery_Date = Convert.ToDateTime(dr["Load End Date/Expected Delivery Date"]);
                        _newRow["Delivery_Date"] = Delivery_Date.ToString("MM/dd/yyyy");
                        _newRow["Billing_Customer_Number"] = Convert.ToString(dtresultCustomerNbrMapping.Rows[0]["BillingCustomerNumber"]);
                        _newRow["Customer_Reference"] = Convert.ToString(dr["Load ID"]);
                        _newRow["ServiceTypeValue"] = Convert.ToString(dtConfiguredData.Rows[0]["ServiceTypeValue"]);
                        _newRow["EnteredByValue"] = Convert.ToString(dtConfiguredData.Rows[0]["EntredByValue"]);
                        _newRow["RequestedByValue"] = Convert.ToString(dtConfiguredData.Rows[0]["RequestedByValue"]);
                        _newRow["PickupDeliveryTransferFlagValue"] = Convert.ToString(dtConfiguredData.Rows[0]["PickupDeliveryTransferFlagValue"]);
                        _newRow["pickup_name"] = Convert.ToString(dtresultRouteNbrMapping.Rows[0]["PickName_DeliveyName"]);
                        _newRow["delivery_name"] = Convert.ToString(dtresultRouteNbrMapping.Rows[0]["PickName_DeliveyName"]);
                        dtableOrderTemplate.Rows.Add(_newRow);
                    }
                    else
                    {
                        strExecutionLogMessage = "TGT Customer NBR/Route NBR not found for this Customer Reference " + System.Environment.NewLine;
                        strExecutionLogMessage += "Customer Reference -" + strCustomer_Reference + System.Environment.NewLine;
                        strExecutionLogMessage += "Type -" + type + System.Environment.NewLine;
                        strExecutionLogMessage += "Origin Address -" + strOriginAddress + System.Environment.NewLine;
                        strExecutionLogMessage += "Destination Address -" + strDestinationAddress + System.Environment.NewLine;
                        objCommon.WriteExecutionLog(strExecutionLogMessage);
                    }

                }
                output.Tables.Add(dtableOrderTemplate);
            }
            else
            {
                strExecutionLogMessage = "No Data found after export for this file" + System.Environment.NewLine;
                strExecutionLogMessage = "TGT Customer NBR/Route NBR not found for this file path  " + TgtCustomerMappingFilepath + System.Environment.NewLine;
                string strEmailSubject = "No Data found after export for this file";
                objCommon.SendExceptionMail(strEmailSubject, strExecutionLogMessage);
                objCommon.WriteExecutionLog(strExecutionLogMessage);
            }
            return output;
        }

        private static DataSet LASGenerateOrderDataTable(DataTable dtinputData, DataTable dtConfiguredData, string type)
        {

            clsCommon objCommon = new clsCommon();
            DataSet output = new DataSet();
            dtinputData.Columns.Add("Billing_Customer_Number");
            dtinputData.Columns.Add("ServiceTypeValue");
            dtinputData.Columns.Add("EnteredByValue");
            dtinputData.Columns.Add("RequestedByValue");
            dtinputData.Columns.Add("PickupDeliveryTransferFlagValue");
            dtinputData.Columns.Add("pickup_name_Value");
            dtinputData.Columns.Add("Pickup_will_be_ready_by_Value");
            dtinputData.Columns.Add("Pickup_no_later_than_Value");
            dtinputData.Columns.Add("Pickup_actual_arrival_time_Value");
            dtinputData.Columns.Add("Pickup_actual_depart_time_Value");
            dtinputData.Columns.Add("Deliver_no_earlier_than_Value");
            dtinputData.Columns.Add("Deliver_no_later_than_Value");
            dtinputData.Columns.Add("Delivery_actual_arrive_time_Value");
            dtinputData.Columns.Add("Delivery_actual_depart_time_Value");
            dtinputData.Columns.Add("CustomerName_Value");
            dtinputData.Columns.Add("Correct_Driver_Number_Value");

            dtinputData.Columns.Add("AddressValue");
            dtinputData.Columns.Add("CityValue");
            dtinputData.Columns.Add("StateValue");
            dtinputData.Columns.Add("ZipValue");
            dtinputData.Columns.Add("Delivery_text_signature_Value");
            // to remove row from data table 
            foreach (DataRow dr in dtinputData.Rows)
            {
                object value = dr["Invoice Date (date activity took place)"];
                if (value == DBNull.Value)
                {
                    dr.Delete();
                }
            }
            dtinputData.AcceptChanges();

            foreach (DataRow dr in dtinputData.Rows)
            {
                dr["Billing_Customer_Number"] = Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                dr["ServiceTypeValue"] = Convert.ToString(dtConfiguredData.Rows[0]["ServiceTypeValue"]);
                dr["EnteredByValue"] = Convert.ToString(dtConfiguredData.Rows[0]["EntredByValue"]);
                dr["RequestedByValue"] = Convert.ToString(dtConfiguredData.Rows[0]["RequestedByValue"]);
                dr["PickupDeliveryTransferFlagValue"] = Convert.ToString(dtConfiguredData.Rows[0]["PickupDeliveryTransferFlagValue"]);
                dr["pickup_name_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["pickup_name_Value"]);
                dr["Pickup_will_be_ready_by_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_will_be_ready_by_Value"]);
                dr["Pickup_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_no_later_than_Value"]);
                dr["Pickup_actual_arrival_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_arrival_time_Value"]);
                dr["Pickup_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_depart_time_Value"]);
                dr["Deliver_no_earlier_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_earlier_than_Value"]);
                dr["Deliver_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_later_than_Value"]);
                dr["Delivery_actual_arrive_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_actual_arrive_time_Value"]);
                dr["Delivery_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_actual_depart_time_Value"]);
                dr["CustomerName_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["CustomerName_Value"]);
                dr["Correct_Driver_Number_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Correct_Driver_Number_Value"]);
                dr["AddressValue"] = Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                dr["CityValue"] = Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]);
                dr["StateValue"] = Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                dr["ZipValue"] = Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);
                dr["Delivery_text_signature_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_text_signature_Value"]);

            }
            DataTable dtCopy = dtinputData.Copy();
            output.Tables.Add(dtCopy);
            return output;
        }

        private static DataSet CDSGenerateOrderDataTable(DataTable dtinputData, DataTable dtConfiguredData, string type)
        {

            //DataView view = dtinputData.DefaultView;
            //view.Sort = "Actual Delivery Date,CDS# ASC";
            //dtinputData = view.ToTable();

            dtinputData = dtinputData.AsEnumerable()
                .OrderBy(en => en.Field<string>("Actual Delivery Date"))
                .ThenBy(en => en.Field<string>("CDS#")).CopyToDataTable();

            //DataView view1 = dtinputData.DefaultView;
            //view1.Sort = "CDS# ASC";
            //dtinputData = view1.ToTable();

            clsCommon objCommon = new clsCommon();
            DataSet output = new DataSet();
            dtinputData.Columns.Add("Billing_Customer_Number");
            dtinputData.Columns.Add("ServiceTypeValue");
            dtinputData.Columns.Add("EnteredByValue");
            dtinputData.Columns.Add("RequestedByValue");
            dtinputData.Columns.Add("PickupDeliveryTransferFlagValue");
            dtinputData.Columns.Add("pickup_name_Value");
            dtinputData.Columns.Add("Pickup_will_be_ready_by_Value");
            dtinputData.Columns.Add("Pickup_no_later_than_Value");
            dtinputData.Columns.Add("Pickup_actual_arrival_time_Value");
            dtinputData.Columns.Add("Pickup_actual_depart_time_Value");
            dtinputData.Columns.Add("Deliver_no_earlier_than_Value");
            dtinputData.Columns.Add("Deliver_no_later_than_Value");
            dtinputData.Columns.Add("Delivery_actual_arrive_time_Value");
            dtinputData.Columns.Add("Delivery_actual_depart_time_Value");
            // dtinputData.Columns.Add("CustomerName_Value");
            // dtinputData.Columns.Add("Correct_Driver_Number_Value");

            //dtinputData.Columns.Add("AddressValue");
            //dtinputData.Columns.Add("CityValue");
            //dtinputData.Columns.Add("StateValue");
            //dtinputData.Columns.Add("ZipValue");
            dtinputData.Columns.Add("Delivery_text_signature_Value");
            dtinputData.Columns.Add("Pickup_address_Value");
            dtinputData.Columns.Add("Pickup_city_Value");
            dtinputData.Columns.Add("Pickup_state/province_Value");
            dtinputData.Columns.Add("Pickup_zip/postal_code_Value");
            dtinputData.Columns.Add("Pickup_text_signature_Value");

            // to remove row from data table 
            //foreach (DataRow dr in dtinputData.Rows)
            //{
            //    object value = dr["CDS#"];
            //    if (value == DBNull.Value || string.IsNullOrEmpty(Convert.ToString(value)))
            //    {
            //        dr.Delete();
            //    }
            //}

            for (int i = 0; i < dtinputData.Rows.Count; i++)
            {
                var temp = dtinputData.Rows[i][0];
                for (int j = 0; j < dtinputData.Rows.Count; j++)
                {
                    DataRow rows = dtinputData.Rows[j];
                    if (temp == rows[0].ToString() && string.IsNullOrEmpty(rows[0].ToString()))
                    {
                        dtinputData.Rows.Remove(rows);      //Update happen here
                    }
                }
            }
            dtinputData.AcceptChanges();



            foreach (DataRow dr in dtinputData.Rows)
            {
                DateTime deliveryDate = Convert.ToDateTime(dr["Actual Delivery Date"]);
                dr["Actual Delivery Date"] = deliveryDate.ToString("MM/dd/yyyy");
                dr["Billing_Customer_Number"] = Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                dr["ServiceTypeValue"] = Convert.ToString(dtConfiguredData.Rows[0]["ServiceTypeValue"]);
                dr["EnteredByValue"] = Convert.ToString(dtConfiguredData.Rows[0]["EntredByValue"]);
                dr["RequestedByValue"] = Convert.ToString(dtConfiguredData.Rows[0]["RequestedByValue"]);
                dr["PickupDeliveryTransferFlagValue"] = Convert.ToString(dtConfiguredData.Rows[0]["PickupDeliveryTransferFlagValue"]);
                dr["pickup_name_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["pickup_name_Value"]);
                dr["Pickup_will_be_ready_by_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_will_be_ready_by_Value"]);
                dr["Pickup_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_no_later_than_Value"]);
                dr["Pickup_actual_arrival_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_arrival_time_Value"]);
                dr["Pickup_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_depart_time_Value"]);
                dr["Deliver_no_earlier_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_earlier_than_Value"]);
                dr["Deliver_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_later_than_Value"]);
                // dr["CustomerName_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["CustomerName_Value"]);
                //  dr["Correct_Driver_Number_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Correct_Driver_Number_Value"]);
                // dr["AddressValue"] = Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                // dr["CityValue"] = Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]);
                // dr["StateValue"] = Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                // dr["ZipValue"] = Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);
                dr["Delivery_text_signature_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_text_signature_Value"]);
                dr["Delivery_actual_arrive_time_Value"] = deliveryDate.ToString("HH:mm:ss tt");
                int AddIntMin = Convert.ToInt32(objCommon.GetConfigValue("CDS_Deliveryactualdeparttime_AddMin"));
                DateTime deliveryDatedepart_time = deliveryDate.AddMinutes(AddIntMin);
                dr["Delivery_actual_depart_time_Value"] = deliveryDatedepart_time.ToString("HH:mm:ss tt");

                dr["Pickup_address_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_address_Value"]);
                dr["Pickup_city_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_city_Value"]);
                dr["Pickup_state/province_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_state/province_Value"]);
                dr["Pickup_zip/postal_code_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_zip/postal_code_Value"]);
                dr["Pickup_text_signature_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_text_signature_Value"]);


            }
            DataTable dtCopy = dtinputData.Copy();
            output.Tables.Add(dtCopy);
            return output;
        }



        private static DataSet ProcessBBBDELfileToDataTable()
        {


            DataSet ds = new DataSet();
            clsCommon objCommon = new clsCommon();
            int company_no = 1;
            int customerNumber = 1605;
            int weight = 7000;
            string storenumber = "61";
            int band = 0;
            double minRate = 0;
            double maxRate = 0;
            double actualRate = 0;
            int rangeSerial = 0;
            double billRate = 0;
            double billRate1 = 0;

            DataTable dtBBBStoreBands = new DataTable();
            DataTable dtBBBDificitWeightRating = new DataTable();
            clsCommon.DSResponse objfscRatesResponse = new clsCommon.DSResponse();
            objfscRatesResponse = objCommon.GetStoreBand_DeficitWeightRatingDetails(Convert.ToInt32(company_no), Convert.ToInt32(customerNumber));
            if (objfscRatesResponse.dsResp.ResponseVal)
            {
                if (objfscRatesResponse.DS.Tables.Count > 0)
                {
                    dtBBBStoreBands = objfscRatesResponse.DS.Tables[0];
                }
                if (objfscRatesResponse.DS.Tables.Count > 1)
                {
                    dtBBBDificitWeightRating = objfscRatesResponse.DS.Tables[1];
                }
            }

            // Foreach will start from here

            DataTable dtstorebandsfiltered = new DataTable();

            IEnumerable<DataRow> storebandsfilteredRows = dtBBBStoreBands.AsEnumerable()
                                                  .Where(row => row.Field<string>("Store") == storenumber);

            if (storebandsfilteredRows.Any())
            {
                dtstorebandsfiltered = storebandsfilteredRows.CopyToDataTable();
            }

            if (dtstorebandsfiltered.Rows.Count > 0)
            {
                band = Convert.ToInt16(dtstorebandsfiltered.Rows[0]["Band"]);
            }


            DataTable dtDeficitWeightRatingfiltered = new DataTable();
            IEnumerable<DataRow> deficitWeightRatingfilteredRows = dtBBBDificitWeightRating.AsEnumerable()
                                                                      .Where(row => (row.Field<int>("Band") == band)
                                                                      && (row.Field<int>("WeightFrom") <= weight)
                                                                      && (weight <= row.Field<int>("WeightTo")));

            if (deficitWeightRatingfilteredRows.Any())
            {
                dtDeficitWeightRatingfiltered = deficitWeightRatingfilteredRows.CopyToDataTable();
            }

            if (dtDeficitWeightRatingfiltered.Rows.Count > 0)
            {

                actualRate = Convert.ToDouble(dtDeficitWeightRatingfiltered.Rows[0]["Rate"]);
                minRate = Convert.ToDouble(dtDeficitWeightRatingfiltered.Rows[0]["MinRate"]);
                maxRate = Convert.ToDouble(dtDeficitWeightRatingfiltered.Rows[0]["MaxRate"]);
                rangeSerial = Convert.ToInt32(dtDeficitWeightRatingfiltered.Rows[0]["RangeSerial"]);

                billRate = (weight / 100) * actualRate;

                if (billRate < minRate)
                {
                    billRate = minRate;
                }
                if (billRate > minRate)
                {
                    rangeSerial = rangeSerial + 1;

                    DataTable dtDeficitWeightRatingfiltered1 = new DataTable();
                    IEnumerable<DataRow> deficitWeightRatingfilteredRows1 = dtBBBDificitWeightRating.AsEnumerable()
                                                                              .Where(row => (row.Field<int>("Band") == band)
                                                                              && (row.Field<int>("RangeSerial") == rangeSerial));
                    if (deficitWeightRatingfilteredRows1.Any())
                    {
                        dtDeficitWeightRatingfiltered = deficitWeightRatingfilteredRows1.CopyToDataTable();
                        int weightFrom = Convert.ToInt32(dtDeficitWeightRatingfiltered.Rows[0]["WeightFrom"]);
                        int WeightTo = Convert.ToInt32(dtDeficitWeightRatingfiltered.Rows[0]["WeightTo"]);
                        actualRate = Convert.ToDouble(dtDeficitWeightRatingfiltered.Rows[0]["Rate"]);
                        billRate1 = (weightFrom / 100) * actualRate;

                        if (billRate1 < billRate)
                        {
                            billRate = billRate1;
                        }
                    }
                }

                if (billRate > maxRate)
                {
                    billRate = maxRate;
                }
                Console.WriteLine("billRate", billRate);

            }

            return ds;
        }

        private static DataSet ConvertBBBTNDFlatfileToDataTable(string filePath, DataTable dtConfiguredData)
        {
            //DataTable tbl = new DataTable();
            DataSet output = new DataSet();
            DataTable tbl = new DataTable();
            tbl.Clear();

            tbl.Columns.Add("PCS_MAN");
            tbl.Columns.Add("CLOSE_DT");
            tbl.Columns.Add("BBB_MAN");
            tbl.Columns.Add("CART");
            tbl.Columns.Add("WEIGHT");
            tbl.Columns.Add("CUBE");
            tbl.Columns.Add("CLASS");
            tbl.Columns.Add("ORIG_SHIP_ID");
            tbl.Columns.Add("SHIP_TYPE");
            tbl.Columns.Add("SHIP_ID");
            tbl.Columns.Add("SHIP_NAME");
            tbl.Columns.Add("SHIP_ADDR_1");
            tbl.Columns.Add("SHIP_ADDR_2");
            tbl.Columns.Add("SHIP_ADDR_3");
            tbl.Columns.Add("SHIP_CITY");
            tbl.Columns.Add("SHIP_STATE");
            tbl.Columns.Add("SHIP_ZIP");
            tbl.Columns.Add("Billing_Customer_Number");
            tbl.Columns.Add("Service_Type");
            tbl.Columns.Add("Entered_by");
            tbl.Columns.Add("Pickup_Delivery_Transfer_Flag");
            tbl.Columns.Add("dim_height");
            tbl.Columns.Add("dim_length");
            tbl.Columns.Add("dim_width");
            tbl.Columns.Add("Pickup_address");
            tbl.Columns.Add("Pickup_city");
            tbl.Columns.Add("Pickup_state");
            tbl.Columns.Add("Pickup_zip");
            tbl.Columns.Add("pickup_name");
            tbl.Columns.Add("pickup_attention");
            tbl.Columns.Add("deliver_attention");
            tbl.Columns.Add("delivery_name");

            tbl.Columns.Add("AddressValue");
            tbl.Columns.Add("CityValue");
            tbl.Columns.Add("StateValue");
            tbl.Columns.Add("ZipValue");

            tbl.Columns.Add("pickup_name_Value");
            tbl.Columns.Add("Pickup_will_be_ready_by_Value");
            tbl.Columns.Add("Pickup_no_later_than_Value");
            tbl.Columns.Add("Pickup_actual_arrival_time_Value");
            tbl.Columns.Add("Pickup_actual_depart_time_Value");
            tbl.Columns.Add("Deliver_no_earlier_than_Value");
            tbl.Columns.Add("Deliver_no_later_than_Value");
            tbl.Columns.Add("Delivery_actual_arrive_time_Value");
            tbl.Columns.Add("Delivery_actual_depart_time_Value");
            tbl.Columns.Add("CustomerName_Value");
            tbl.Columns.Add("Correct_Driver_Number_Value");
            tbl.Columns.Add("Delivery_text_signature_Value");



            string[] lines = System.IO.File.ReadAllLines(filePath);
            lines = lines.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            foreach (string line in lines)
            {
                // Console.WriteLine("line length" + line.Length);
                DataRow _newRow = tbl.NewRow();
                _newRow["PCS_MAN"] = "\t" + line.Substring(0, 6).Trim();
                // _newRow["pickup_name"] = "\t" + line.Substring(0, 6).Trim();
                //_newRow["CLOSE_DT"] = line.Substring(6, 8);
                _newRow["CLOSE_DT"] = "\t" + DateTime.ParseExact(line.Substring(6, 8), "yyyyMMdd",
                          CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");

                string strbbbman = line.Substring(14, 30).Trim();
                // _newRow["BBB_MAN"] = "\t" + line.Substring(14, 30).Trim();
                _newRow["BBB_MAN"] = "\t" + Right(strbbbman, 23);
                _newRow["delivery_name"] = "\t" + strbbbman;
                _newRow["CART"] = "\t" + line.Substring(44, 28).Trim();
                _newRow["WEIGHT"] = "\t" + line.Substring(72, 12).Trim();
                _newRow["CUBE"] = "\t" + line.Substring(84, 12).Trim();
                _newRow["dim_height"] = "\t" + line.Substring(84, 12).Trim();
                _newRow["dim_length"] = "\t" + line.Substring(84, 12).Trim();
                _newRow["dim_width"] = "\t" + line.Substring(84, 12).Trim();
                _newRow["CLASS"] = "\t" + line.Substring(96, 7).Trim();
                _newRow["ORIG_SHIP_ID"] = "\t" + line.Substring(103, 11).Trim();
                _newRow["SHIP_TYPE"] = line.Substring(114, 1).Trim();
                if (!string.IsNullOrEmpty(line.Substring(115, 8).Trim()))
                {
                    _newRow["SHIP_ID"] = "\t" + line.Substring(115, 8).Trim();
                }
                else
                {
                    _newRow["SHIP_ID"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["RequestedByValue"]);
                }

                _newRow["SHIP_NAME"] = "\t" + line.Substring(153, 30).Trim();
                if (!string.IsNullOrEmpty(line.Substring(123, 30).Trim()))
                {
                    _newRow["SHIP_ADDR_1"] = "\t" + line.Substring(123, 30).Trim();
                }
                else
                {
                    _newRow["SHIP_ADDR_1"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                }
                _newRow["SHIP_ADDR_2"] = "\t" + line.Substring(183, 30).Trim();
                _newRow["SHIP_ADDR_3"] = "\t" + line.Substring(213, 30).Trim();
                if (!string.IsNullOrEmpty(line.Substring(243, 20).Trim()))
                {
                    _newRow["SHIP_CITY"] = "\t" + line.Substring(243, 20).Trim();
                }
                else
                {
                    _newRow["SHIP_CITY"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]);
                }
                if (!string.IsNullOrEmpty(line.Substring(263, 2).Trim()))
                {
                    _newRow["SHIP_STATE"] = "\t" + line.Substring(263, 2).Trim();
                }
                else
                {
                    _newRow["SHIP_STATE"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                }
                if (!string.IsNullOrEmpty(line.Substring(265, 10).Trim()))
                {
                    _newRow["SHIP_ZIP"] = "\t" + line.Substring(265, 10).Trim();
                }
                else
                {
                    _newRow["SHIP_ZIP"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);
                }
                _newRow["Billing_Customer_Number"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                _newRow["Service_Type"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ServiceTypeValue"]);
                _newRow["Entered_by"] = Convert.ToString(dtConfiguredData.Rows[0]["EntredByValue"]);
                _newRow["Pickup_Delivery_Transfer_Flag"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["PickupDeliveryTransferFlagValue"]);
                _newRow["Pickup_address"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                _newRow["Pickup_city"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]); ;
                _newRow["Pickup_state"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                _newRow["Pickup_zip"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);

                _newRow["AddressValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                _newRow["CityValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]);
                _newRow["StateValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                _newRow["ZipValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);

                _newRow["pickup_name_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["pickup_name_Value"]);
                _newRow["Pickup_will_be_ready_by_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_will_be_ready_by_Value"]);
                _newRow["Pickup_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_no_later_than_Value"]);
                _newRow["Pickup_actual_arrival_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_arrival_time_Value"]);
                _newRow["Pickup_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_depart_time_Value"]);
                _newRow["Deliver_no_earlier_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_earlier_than_Value"]);
                _newRow["Deliver_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_later_than_Value"]);
                _newRow["Delivery_actual_arrive_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_actual_arrive_time_Value"]);
                _newRow["Delivery_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_actual_depart_time_Value"]);
                _newRow["CustomerName_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["CustomerName_Value"]);
                _newRow["Correct_Driver_Number_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Correct_Driver_Number_Value"]);
                _newRow["Delivery_text_signature_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_text_signature_Value"]);

                tbl.Rows.Add(_newRow);
            }
            output.Tables.Add(tbl);
            return output;
        }

        private static DataSet ConvertBBBDELFlatfileToDataTable(string filePath, DataTable dtConfiguredData)
        {
            DataSet output = new DataSet();
            DataTable tbl = new DataTable();
            tbl.Clear();

            tbl.Columns.Add("PCS_MAN");
            tbl.Columns.Add("CLOSE_DT");
            tbl.Columns.Add("BBB_MAN");
            tbl.Columns.Add("CART");
            tbl.Columns.Add("WEIGHT");
            tbl.Columns.Add("CUBE");
            tbl.Columns.Add("CLASS");
            tbl.Columns.Add("ORIG_SHIP_ID");
            tbl.Columns.Add("SHIP_TYPE");
            tbl.Columns.Add("SHIP_ID");
            tbl.Columns.Add("SHIP_NAME");
            tbl.Columns.Add("SHIP_ADDR_1");
            tbl.Columns.Add("SHIP_ADDR_2");
            tbl.Columns.Add("SHIP_ADDR_3");
            tbl.Columns.Add("SHIP_CITY");
            tbl.Columns.Add("SHIP_STATE");
            tbl.Columns.Add("SHIP_ZIP");
            tbl.Columns.Add("Billing_Customer_Number");
            tbl.Columns.Add("Service_Type");
            tbl.Columns.Add("Entered_by");
            tbl.Columns.Add("Pickup_Delivery_Transfer_Flag");
            tbl.Columns.Add("dim_height");
            tbl.Columns.Add("dim_length");
            tbl.Columns.Add("dim_width");
            tbl.Columns.Add("Pickup_address");
            tbl.Columns.Add("Pickup_city");
            tbl.Columns.Add("Pickup_state");
            tbl.Columns.Add("Pickup_zip");
            tbl.Columns.Add("pickup_name");
            tbl.Columns.Add("pickup_attention");
            tbl.Columns.Add("deliver_attention");
            tbl.Columns.Add("delivery_name");

            tbl.Columns.Add("AddressValue");
            tbl.Columns.Add("CityValue");
            tbl.Columns.Add("StateValue");
            tbl.Columns.Add("ZipValue");

            tbl.Columns.Add("pickup_name_Value");
            tbl.Columns.Add("Pickup_will_be_ready_by_Value");
            tbl.Columns.Add("Pickup_no_later_than_Value");
            tbl.Columns.Add("Pickup_actual_arrival_time_Value");
            tbl.Columns.Add("Pickup_actual_depart_time_Value");
            tbl.Columns.Add("Deliver_no_earlier_than_Value");
            tbl.Columns.Add("Deliver_no_later_than_Value");
            tbl.Columns.Add("Delivery_actual_arrive_time_Value");
            tbl.Columns.Add("Delivery_actual_depart_time_Value");
            tbl.Columns.Add("CustomerName_Value");
            tbl.Columns.Add("Correct_Driver_Number_Value");
            tbl.Columns.Add("Delivery_text_signature_Value");
            // tbl.Columns.Add("Store");


            string[] lines = System.IO.File.ReadAllLines(filePath);
            lines = lines.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            foreach (string line in lines)
            {
                DataRow _newRow = tbl.NewRow();
                _newRow["PCS_MAN"] = "\t" + line.Substring(0, 6).Trim();
                _newRow["CLOSE_DT"] = "\t" + DateTime.ParseExact(line.Substring(6, 8), "yyyyMMdd",
                          CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");

                string strbbbman = line.Substring(14, 30).Trim();
                _newRow["BBB_MAN"] = Right(strbbbman, 23);
                //_newRow["Store"] = strbbbman.Substring(18, 22).Trim();
                _newRow["delivery_name"] = "\t" + strbbbman;
                _newRow["CART"] = "\t" + line.Substring(44, 28).Trim();
                _newRow["WEIGHT"] = line.Substring(72, 12).Trim();
                _newRow["CUBE"] = "\t" + line.Substring(84, 12).Trim();
                _newRow["dim_height"] = "\t" + line.Substring(84, 12).Trim();
                _newRow["dim_length"] = "\t" + line.Substring(84, 12).Trim();
                _newRow["dim_width"] = "\t" + line.Substring(84, 12).Trim();
                _newRow["CLASS"] = "\t" + line.Substring(96, 7).Trim();
                _newRow["ORIG_SHIP_ID"] = "\t" + line.Substring(103, 11).Trim();
                _newRow["SHIP_TYPE"] = line.Substring(114, 1).Trim();
                if (!string.IsNullOrEmpty(line.Substring(115, 8).Trim()))
                {
                    _newRow["SHIP_ID"] = "\t" + line.Substring(115, 8).Trim();
                }
                else
                {
                    _newRow["SHIP_ID"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["RequestedByValue"]);
                }

                _newRow["SHIP_NAME"] = "\t" + line.Substring(153, 30).Trim();
                if (!string.IsNullOrEmpty(line.Substring(123, 30).Trim()))
                {
                    _newRow["SHIP_ADDR_1"] = "\t" + line.Substring(123, 30).Trim();
                }
                else
                {
                    _newRow["SHIP_ADDR_1"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                }
                _newRow["SHIP_ADDR_2"] = "\t" + line.Substring(183, 30).Trim();
                _newRow["SHIP_ADDR_3"] = "\t" + line.Substring(213, 30).Trim();
                if (!string.IsNullOrEmpty(line.Substring(243, 20).Trim()))
                {
                    _newRow["SHIP_CITY"] = "\t" + line.Substring(243, 20).Trim();
                }
                else
                {
                    _newRow["SHIP_CITY"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]);
                }
                if (!string.IsNullOrEmpty(line.Substring(263, 2).Trim()))
                {
                    _newRow["SHIP_STATE"] = "\t" + line.Substring(263, 2).Trim();
                }
                else
                {
                    _newRow["SHIP_STATE"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                }
                if (!string.IsNullOrEmpty(line.Substring(265, 10).Trim()))
                {
                    _newRow["SHIP_ZIP"] = "\t" + line.Substring(265, 10).Trim();
                }
                else
                {
                    _newRow["SHIP_ZIP"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);
                }
                _newRow["Billing_Customer_Number"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                _newRow["Service_Type"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ServiceTypeValue"]);
                _newRow["Entered_by"] = Convert.ToString(dtConfiguredData.Rows[0]["EntredByValue"]);
                _newRow["Pickup_Delivery_Transfer_Flag"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["PickupDeliveryTransferFlagValue"]);
                _newRow["Pickup_address"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                _newRow["Pickup_city"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]); ;
                _newRow["Pickup_state"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                _newRow["Pickup_zip"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);

                _newRow["AddressValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["AddressValue"]);
                _newRow["CityValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["CityValue"]);
                _newRow["StateValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["StateValue"]);
                _newRow["ZipValue"] = "\t" + Convert.ToString(dtConfiguredData.Rows[0]["ZipValue"]);

                _newRow["pickup_name_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["pickup_name_Value"]);
                _newRow["Pickup_will_be_ready_by_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_will_be_ready_by_Value"]);
                _newRow["Pickup_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_no_later_than_Value"]);
                _newRow["Pickup_actual_arrival_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_arrival_time_Value"]);
                _newRow["Pickup_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Pickup_actual_depart_time_Value"]);
                _newRow["Deliver_no_earlier_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_earlier_than_Value"]);
                _newRow["Deliver_no_later_than_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Deliver_no_later_than_Value"]);
                _newRow["Delivery_actual_arrive_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_actual_arrive_time_Value"]);
                _newRow["Delivery_actual_depart_time_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_actual_depart_time_Value"]);
                _newRow["CustomerName_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["CustomerName_Value"]);
                _newRow["Correct_Driver_Number_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Correct_Driver_Number_Value"]);
                _newRow["Delivery_text_signature_Value"] = Convert.ToString(dtConfiguredData.Rows[0]["Delivery_text_signature_Value"]);

                tbl.Rows.Add(_newRow);
            }
            output.Tables.Add(tbl);
            return output;
        }
        private static void ProcessBBBDELfileToDataTable(string fileName, DataTable dtableOrderTemplate,
            int company_no, int customerNumber, string attachmentPath)
        {
            DataSet ds = new DataSet();
            clsCommon objCommon = new clsCommon();
            int band = 0;
            double minRate = 0;
            double maxRate = 0;
            double actualRate = 0;
            int rangeSerial = 0;
            double billRate = 0;
            double billRate1 = 0;

            string executionLogMessage = "";
            string OriginlafileName = fileName;
            try
            {

                fileName = fileName.Replace("TND", "DEL");

                DataTable dtableOrderTemplateFinal = new DataTable();
                // to filter the data based on the customer reference and finding the number of pieces 

                //int BBBEnableCartoncountsummaryOnly_orderlineitemAbove = int.Parse(objCommon.GetConfigValue("BBBEnableCartoncountsummaryOnly_orderlineitemAbove"));
                //if (dtableOrderTemplate.Rows.Count < BBBEnableCartoncountsummaryOnly_orderlineitemAbove)
                //{
                //    dtableOrderTemplateFinal = dtableOrderTemplate.Copy();
                //}
                //else
                //{
                DataView view = new DataView(dtableOrderTemplate);
                DataTable dtdistinctValues = view.ToTable(true, "Customer Reference");
                foreach (DataRow dr in dtdistinctValues.Rows)
                {
                    object value = dr["Customer Reference"];
                    if (value == DBNull.Value)
                        break;
                    string ReferenceId = Convert.ToString(dr["Customer Reference"]);
                    try
                    {
                        DataTable dtWeightData = new DataTable();
                        dtWeightData.Columns.Add("Customer Reference", typeof(string));
                        dtWeightData.Columns.Add("Weight", typeof(int));

                        if (dtableOrderTemplateFinal.Rows.Count > 0)
                        {

                            DataTable dtresult = dtableOrderTemplate.Select("[Customer Reference]= '" + dr["Customer Reference"] + "'").CopyToDataTable();
                            for (int row = 0; row < dtresult.Rows.Count;)
                            {
                                DataRow dr1 = dtableOrderTemplateFinal.NewRow();
                                for (int column = 0; column < dtresult.Columns.Count; column++)
                                {
                                    dr1[column] = dtresult.Rows[row][column];
                                }
                                dr1["Pieces"] = dtresult.Rows.Count;

                                //try
                                //{
                                dtWeightData.Clear();
                                dtWeightData = dtresult.AsEnumerable()
                                     .GroupBy(r => r.Field<string>("Customer Reference"))
                                     .Select(g =>
                                     {
                                         var row1 = dtWeightData.NewRow();
                                         row1["Customer Reference"] = g.Key;
                                         row1["Weight"] = g.Sum(r => Convert.ToDouble(r.Field<string>("Weight")));
                                         return row1;
                                     }).CopyToDataTable();

                                dr1["Weight"] = dtWeightData.Rows[0]["Weight"];
                                //}
                                //catch (Exception ex)
                                //{
                                //    executionLogMessage = "Exception found for this file: " + fileName + System.Environment.NewLine;
                                //    executionLogMessage += "while summarisation of the record";
                                //    executionLogMessage += ex.Message;
                                //    objCommon.WriteExecutionLog(executionLogMessage);
                                //    objCommon.WriteErrorLog(ex, "ProcessBBBDELfileToDataTable");
                                //}
                                dtableOrderTemplateFinal.Rows.Add(dr1.ItemArray);
                                break;
                            }
                        }
                        else
                        {
                            DataRow[] drresult = dtableOrderTemplate.Select("[Customer Reference]= '" + dr["Customer Reference"] + "'");
                            var firstRow = drresult.AsEnumerable().First();
                            dtableOrderTemplateFinal = new[] { firstRow }.CopyToDataTable();
                            dtableOrderTemplateFinal.Rows[0]["Pieces"] = drresult.Length;

                            DataTable dtTemptableOrderTemplate = new DataTable();

                            //try
                            //{
                            string customerrefernce = Convert.ToString(dr["Customer Reference"]);
                            dtTemptableOrderTemplate = dtableOrderTemplate.AsEnumerable()
                                                                 .Where(x => x.Field<string>("Customer Reference") == customerrefernce)
                                                                 .CopyToDataTable();
                            dtWeightData.Clear();
                            //DataTable dtWeightData = new DataTable();
                            //dtWeightData.Columns.Add("Customer Reference", typeof(string));
                            //dtWeightData.Columns.Add("Weight", typeof(int));

                            dtWeightData = dtTemptableOrderTemplate.AsEnumerable()
                                 .GroupBy(r => r.Field<string>("Customer Reference"))
                                 .Select(g =>
                                 {
                                     var row = dtWeightData.NewRow();
                                     row["Customer Reference"] = g.Key;
                                     row["Weight"] = g.Sum(r => Convert.ToDouble(r.Field<string>("Weight")));
                                     return row;
                                 }).CopyToDataTable();

                            dtableOrderTemplateFinal.Rows[0]["Weight"] = dtWeightData.Rows[0]["Weight"];

                            //}
                            //catch (Exception ex)
                            //{
                            //    executionLogMessage = "Exception found for this file: " + fileName + System.Environment.NewLine;
                            //    executionLogMessage += "while summarisation of the record";
                            //    executionLogMessage += ex.Message;
                            //    objCommon.WriteExecutionLog(executionLogMessage);
                            //    objCommon.WriteErrorLog(ex, "ProcessBBBDELfileToDataTable");
                            //}

                            dtableOrderTemplateFinal.AcceptChanges();
                        }
                    }
                    catch (Exception ex)
                    {
                        executionLogMessage = "BBB summary file Creation Exception -" + ex.Message + System.Environment.NewLine;
                        executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                        executionLogMessage += "Originla file Name " + OriginlafileName;
                        objCommon.WriteErrorLog(ex, executionLogMessage);
                        string subject = "Found Exception for while generating the summary and not processed this file for Delivey, Please look into it immediately, File Name   : " + fileName;
                        objCommon.SendExceptionMail(subject, executionLogMessage);
                    }
                }
                //  }


                DataTable dtBBBStoreBands = new DataTable();
                DataTable dtBBBDificitWeightRating = new DataTable();
                clsCommon.DSResponse objDificitRatesResponse = new clsCommon.DSResponse();
                objDificitRatesResponse = objCommon.GetStoreBand_DeficitWeightRatingDetails(Convert.ToInt32(company_no), Convert.ToInt32(customerNumber));
                if (objDificitRatesResponse.dsResp.ResponseVal)
                {
                    if (objDificitRatesResponse.DS.Tables.Count > 0)
                    {
                        dtBBBStoreBands = objDificitRatesResponse.DS.Tables[0];
                    }
                    if (objDificitRatesResponse.DS.Tables.Count > 1)
                    {
                        dtBBBDificitWeightRating = objDificitRatesResponse.DS.Tables[1];
                    }
                }

                // Foreach will start from here

                DataTable dtBillingRates = new DataTable();
                DataTable dtPayableRates = new DataTable();
                clsCommon.DSResponse objBPRatesResponse = new clsCommon.DSResponse();
                objBPRatesResponse = objCommon.GetBillingRatesAndPayableRates_CustomerMappingDetails(company_no.ToString(), customerNumber.ToString());
                if (objBPRatesResponse.dsResp.ResponseVal)
                {

                    if (objBPRatesResponse.DS.Tables.Count > 0)
                    {
                        dtBillingRates = objBPRatesResponse.DS.Tables[0].Copy();
                    }

                    if (objBPRatesResponse.DS.Tables.Count > 1)
                    {
                        dtPayableRates = objBPRatesResponse.DS.Tables[1].Copy();
                    }
                }

                foreach (DataRow dr in dtableOrderTemplateFinal.Rows)
                {
                    object value = dr["Delivery Date"];
                    if (value == DBNull.Value)
                        break;


                    band = 0;
                    minRate = 0;
                    maxRate = 0;
                    actualRate = 0;
                    rangeSerial = 0;
                    billRate = 0;
                    billRate1 = 0;

                    int weight = Convert.ToInt32((dr["Weight"]));
                    string bolNumber = (Convert.ToString(dr["BOL Number"]));

                    string storenumber = Regex.Replace(bolNumber, @"\t", "");

                    //if(storenumber.Length > 6)
                    //{
                    //    storenumber = storenumber.Substring(3, 4);
                    //}

                    if (IsDigitsOnly(storenumber))
                    {
                        storenumber = Convert.ToString(Convert.ToInt32(storenumber));
                    }

                    DataTable dtstorebandsfiltered = new DataTable();

                    IEnumerable<DataRow> storebandsfilteredRows = dtBBBStoreBands.AsEnumerable()
                                                          .Where(row => row.Field<string>("Store") == storenumber);
                    if (storebandsfilteredRows.Any())
                    {
                        dtstorebandsfiltered = storebandsfilteredRows.CopyToDataTable();
                    }

                    if (dtstorebandsfiltered.Rows.Count > 0)
                    {
                        band = Convert.ToInt16(dtstorebandsfiltered.Rows[0]["Band"]);
                    }
                    else
                    {
                        executionLogMessage = "Band Details not found for this store number : " + storenumber + System.Environment.NewLine;
                        executionLogMessage += "So not able to process this file, please update the Store Band details table  with appropriate values." + System.Environment.NewLine;
                        executionLogMessage += "Processing file Name : " + fileName + System.Environment.NewLine;
                        executionLogMessage += "Original file Name : " + OriginlafileName + System.Environment.NewLine;
                        executionLogMessage += "This email as marked, please mark as unread to reprocess process. " + System.Environment.NewLine;

                        objCommon.WriteExecutionLog(executionLogMessage);
                        string fromMail = objCommon.GetConfigValue("FromMailID");
                        string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                        string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                        string toMail = objCommon.GetConfigValue("BBBSendExceptionEmail");
                        string subject = "Band Details not found for this store number : " + storenumber;
                        objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                        throw new NullReferenceException("Band Details not found for this store number : " + storenumber);

                    }

                    DataTable dtDeficitWeightRatingfiltered = new DataTable();
                    IEnumerable<DataRow> deficitWeightRatingfilteredRows = dtBBBDificitWeightRating.AsEnumerable()
                                                                              .Where(row => (row.Field<int>("Band") == band)
                                                                              && (row.Field<int>("WeightFrom") <= weight)
                                                                              && (weight <= row.Field<int>("WeightTo")));

                    if (deficitWeightRatingfilteredRows.Any())
                    {
                        dtDeficitWeightRatingfiltered = deficitWeightRatingfilteredRows.CopyToDataTable();
                    }

                    if (dtDeficitWeightRatingfiltered.Rows.Count > 0)
                    {
                        actualRate = Convert.ToDouble(dtDeficitWeightRatingfiltered.Rows[0]["Rate"]);
                        minRate = Convert.ToDouble(dtDeficitWeightRatingfiltered.Rows[0]["MinRate"]);
                        maxRate = Convert.ToDouble(dtDeficitWeightRatingfiltered.Rows[0]["MaxRate"]);
                        rangeSerial = Convert.ToInt32(dtDeficitWeightRatingfiltered.Rows[0]["RangeSerial"]);

                        billRate = Convert.ToDouble(Convert.ToDouble(Convert.ToDouble(weight) / 100) * actualRate);

                        if (billRate < minRate)
                        {
                            billRate = minRate;
                        }
                        if (billRate > minRate)
                        {
                            rangeSerial = rangeSerial + 1;

                            DataTable dtDeficitWeightRatingfiltered1 = new DataTable();
                            IEnumerable<DataRow> deficitWeightRatingfilteredRows1 = dtBBBDificitWeightRating.AsEnumerable()
                                                                                      .Where(row => (row.Field<int>("Band") == band)
                                                                                      && (row.Field<int>("RangeSerial") == rangeSerial));
                            if (deficitWeightRatingfilteredRows1.Any())
                            {
                                dtDeficitWeightRatingfiltered = deficitWeightRatingfilteredRows1.CopyToDataTable();
                                int weightFrom = Convert.ToInt32(dtDeficitWeightRatingfiltered.Rows[0]["WeightFrom"]);
                                int WeightTo = Convert.ToInt32(dtDeficitWeightRatingfiltered.Rows[0]["WeightTo"]);
                                actualRate = Convert.ToDouble(dtDeficitWeightRatingfiltered.Rows[0]["Rate"]);
                                billRate1 = Convert.ToDouble(Convert.ToDouble(Convert.ToDouble(weightFrom) / 100) * actualRate);

                                if (billRate1 < billRate)
                                {
                                    billRate = billRate1;
                                }
                            }
                        }

                        if (billRate > maxRate)
                        {
                            billRate = maxRate;
                        }

                        dr["Bill Rate"] = Math.Round(billRate);
                    }

                    DateTime dtdeliveryDate = Convert.ToDateTime(Regex.Replace(value.ToString(), @"\t", ""));

                    var invCulture = System.Globalization.CultureInfo.InvariantCulture;

                    DataTable tblBillRatesFiltered = new DataTable();
                    IEnumerable<DataRow> billratesfilteredRows = dtBillingRates.AsEnumerable()
                    .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate")));

                    if (billratesfilteredRows.Any())
                    {
                        tblBillRatesFiltered = billratesfilteredRows.CopyToDataTable();
                    }

                    DataTable tblPayableRatesFiltered = new DataTable();
                    IEnumerable<DataRow> payableratesfilteredRows = dtPayableRates.AsEnumerable()
                    .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate")));


                    if (payableratesfilteredRows.Any())
                    {
                        tblPayableRatesFiltered = payableratesfilteredRows.CopyToDataTable();
                    }

                    if (tblBillRatesFiltered.Rows.Count > 0)
                    {
                        if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                        {
                            //if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"])))
                            //{
                            //    dr["Bill Rate"] = billRate;
                            //}
                            // dr["Bill Rate"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"])))
                                dr["rate_buck_amt2"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"])))
                                dr["Pieces ACC"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"])))
                                dr["FSC"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"])))
                                dr["rate_buck_amt4"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"])))
                                dr["rate_buck_amt5"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"])))
                                dr["rate_buck_amt6"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"])))
                                dr["rate_buck_amt7"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"])))
                                dr["rate_buck_amt8"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"])))
                                dr["rate_buck_amt9"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"])))
                                dr["rate_buck_amt11"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"]);

                        }
                        else
                        {
                            //if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"])))
                            //{
                            //    dr["Bill Rate"] = billRate;
                            //}
                            //dr["Bill Rate"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"])))
                                dr["rate_buck_amt2"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"])))
                                dr["Pieces ACC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"])))
                                dr["FSC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt10"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"])))
                                dr["rate_buck_amt4"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"])))
                                dr["rate_buck_amt5"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"])))
                                dr["rate_buck_amt6"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"])))
                                dr["rate_buck_amt7"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"])))
                                dr["rate_buck_amt8"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"])))
                                dr["rate_buck_amt9"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"])))
                                dr["rate_buck_amt11"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"]);

                        }
                    }

                    if (tblPayableRatesFiltered.Rows.Count > 0)
                    {
                        if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                        {
                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                                dr["Carrier Base Pay"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge5"])))
                                dr["Carrier ACC"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge5"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge6"])))
                                dr["Carrier FSC"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge6"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge2"])))
                                dr["charge2"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge2"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge3"])))
                                dr["charge3"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge3"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge4"])))
                                dr["charge4"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge4"]);

                        }
                        else
                        {

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                                dr["Carrier Base Pay"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge5"])))
                                dr["Carrier ACC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge5"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge6"])))
                                dr["Carrier FSC"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge6"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge2"])))
                                dr["charge2"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge2"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge3"])))
                                dr["charge3"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge3"]);

                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge4"])))
                                dr["charge4"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge4"]);

                        }
                    }
                }

                dtableOrderTemplateFinal.TableName = "Template";
                clsExcelHelper.ExportDataToXLSX(dtableOrderTemplateFinal, attachmentPath, fileName);
                // objCommon.CleanAttachmentWorkingFolder();
            }
            catch (Exception ex)
            {
                executionLogMessage = "Exception found for this file: " + fileName + System.Environment.NewLine;
                executionLogMessage = "Original file: " + OriginlafileName + System.Environment.NewLine;
                executionLogMessage += ex.Message;
                objCommon.WriteExecutionLog(executionLogMessage);
                objCommon.WriteErrorLog(ex, "ProcessBBBDELfileToDataTable");
                string subject = "Found Exception for while processing this file for Delivey, Please look into it immediately, File Name   : " + fileName;
                objCommon.SendExceptionMail(subject, executionLogMessage);
            }
        }

        public static bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }

            return true;
        }
        private static void GenerateOrderTemplate(ref DataTable dtableOrderTemplate, DataTable dtOrderData,
            string fileName, string customerName, string locationCode,
            string productCode, string productSubCode)
        {

            clsCommon objCommon = new clsCommon();

            string executionLogMessage = string.Empty;
            //  DataTable dtableOrderTemplate = new DataTable();
            dtableOrderTemplate.Clear();

            clsCommon.DSResponse objDsRes = new clsCommon.DSResponse();
            objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);
            if (objDsRes.dsResp.ResponseVal)
            {
                string strDatetime = DateTime.Now.ToString("yyyyMMddHHmmss");
                try
                {
                    // DataTable dtOrderData = dsExcel.Tables[0];

                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_Date"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_Date"])].ColumnName = "Delivery Date";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Billing_Customer_Number"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Billing_Customer_Number"])].ColumnName = "Billing Customer Number";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Customer_Reference"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Customer_Reference"])].ColumnName = "Customer Reference";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["BOL_Number"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["BOL_Number"])].ColumnName = "BOL Number";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Customer_Name"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Customer_Name"])].ColumnName = "Customer Name";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Route_Number"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Route_Number"])].ColumnName = "Route Number";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Original_Driver_No"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Original_Driver_No"])].ColumnName = "Original Driver No";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Correct_Driver_Number"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Correct_Driver_Number"])].ColumnName = "Correct Driver Number";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_Name"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_Name"])].ColumnName = "Carrier Name";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Address"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Address"])].ColumnName = "Address";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["City"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["City"])].ColumnName = "City";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["State"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["State"])].ColumnName = "State";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Zip"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Zip"])].ColumnName = "Zip";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pieces"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pieces"])].ColumnName = "Pieces";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Miles"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Miles"])].ColumnName = "Miles";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_Zip"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_Zip"])].ColumnName = "Delivery Zip";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Zip_Code_Surcharge?"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Zip_Code_Surcharge?"])].ColumnName = "Zip Code Surcharge?";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Store_Code"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Store_Code"])].ColumnName = "Store_Code";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Type_of_Delivery"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Type_of_Delivery"])].ColumnName = "Type of Delivery";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Service_Type"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Service_Type"])].ColumnName = "Service Type";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Bill_Rate"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Bill_Rate"])].ColumnName = "Bill Rate";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pieces_ACC"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pieces_ACC"])].ColumnName = "Pieces ACC";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["FSC"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["FSC"])].ColumnName = "FSC";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Total_Bill"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Total_Bill"])].ColumnName = "Total Bill";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_Base_Pay"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_Base_Pay"])].ColumnName = "Carrier Base Pay";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_ACC"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Carrier_ACC"])].ColumnName = "Carrier ACC";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Side_Notes"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Side_Notes"])].ColumnName = "Side Notes";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_requested_date"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_requested_date"])].ColumnName = "Pickup requested date";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_will_be_ready_by"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_will_be_ready_by"])].ColumnName = "Pickup will be ready by";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_no_later_than"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_no_later_than"])].ColumnName = "Pickup no later than";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_date"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_date"])].ColumnName = "Pickup actual date";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_arrival_time"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_arrival_time"])].ColumnName = "Pickup actual arrival time";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_depart_time"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_actual_depart_time"])].ColumnName = "Pickup actual depart time";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_name"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_name"])].ColumnName = "Pickup name";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_address"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_address"])].ColumnName = "Pickup address";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_city"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_city"])].ColumnName = "Pickup city";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_state/province"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_state/province"])].ColumnName = "Pickup state/province";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_zip/postal_code"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_zip/postal_code"])].ColumnName = "Pickup zip/postal code";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_text_signature"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_text_signature"])].ColumnName = "Pickup text signature";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_requested_date"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_requested_date"])].ColumnName = "Delivery requested date";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Deliver_no_earlier_than"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Deliver_no_earlier_than"])].ColumnName = "Deliver no earlier than";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Deliver_no_later_than"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Deliver_no_later_than"])].ColumnName = "Deliver no later than";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_date"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_date"])].ColumnName = "Delivery actual date";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_arrive_time"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_arrive_time"])].ColumnName = "Delivery actual arrive time";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_depart_time"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_actual_depart_time"])].ColumnName = "Delivery actual depart time";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_text_signature"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Delivery_text_signature"])].ColumnName = "Delivery text signature";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Requested_by"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Requested_by"])].ColumnName = "Requested by";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Entered_by"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Entered_by"])].ColumnName = "Entered by";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_Delivery_Transfer_Flag"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Pickup_Delivery_Transfer_Flag"])].ColumnName = "Pickup Delivery Transfer Flag";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["weight"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["weight"])].ColumnName = "Weight";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["insurance_amount"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["insurance_amount"])].ColumnName = "Insurance Amount";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["master_airway_bill_number"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["master_airway_bill_number"])].ColumnName = "Master airway bill number";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["po_number"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["po_number"])].ColumnName = "PO Number";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["house_airway_bill_number"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["house_airway_bill_number"])].ColumnName = "House airway bill number";
                    }

                    //if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Dimensions"])))
                    //{
                    //    dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["Dimensions"])].ColumnName = "Dimensions";
                    //}
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["item_number"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["item_number"])].ColumnName = "Item Number";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["item_description"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["item_description"])].ColumnName = "Item Description";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_height"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_height"])].ColumnName = "Dim Height";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_length"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_length"])].ColumnName = "Dim Length";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_width"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["dim_width"])].ColumnName = "Dim Width";
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["pickup_room"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["pickup_room"])].ColumnName = "Pickup Room";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["pickup_attention"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["pickup_attention"])].ColumnName = "Pickup Attention";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["deliver_attention"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["deliver_attention"])].ColumnName = "Deliver Attention";
                    }


                    if (dtOrderData.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtOrderData.Rows)
                        {
                            DataRow _newRow = dtableOrderTemplate.NewRow();
                            if (dr.Table.Columns.Contains("Delivery Date"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery Date"])))
                                {
                                    _newRow["Delivery Date"] = Convert.ToString(dr["Delivery Date"]);
                                }
                                else
                                {
                                    _newRow["Delivery Date"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Delivery Date"] = "";
                            }

                            _newRow["Company"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["CompanyNumber"]);

                            if (dr.Table.Columns.Contains("Billing Customer Number"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Billing Customer Number"])))
                                {
                                    _newRow["Billing Customer Number"] = Convert.ToString(dr["Billing Customer Number"]);
                                }
                                else
                                {
                                    _newRow["Billing Customer Number"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Billing Customer Number"] = "";
                            }

                            if (dr.Table.Columns.Contains("Customer Reference"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Customer Reference"])))
                                {
                                    _newRow["Customer Reference"] = Convert.ToString(dr["Customer Reference"]);
                                }
                                else
                                {
                                    _newRow["Customer Reference"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Customer Reference"] = "";
                            }
                            if (dr.Table.Columns.Contains("BOL Number"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["BOL Number"])))
                                {
                                    _newRow["BOL Number"] = Convert.ToString(dr["BOL Number"]);
                                }
                                else
                                {
                                    _newRow["BOL Number"] = "";
                                }
                            }
                            else
                            {
                                _newRow["BOL Number"] = "";
                            }
                            if (dr.Table.Columns.Contains("Customer Name"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Customer Name"])))
                                {
                                    _newRow["Customer Name"] = Convert.ToString(dr["Customer Name"]);
                                }
                                else
                                {
                                    _newRow["Customer Name"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Customer Name"] = "";
                            }
                            if (dr.Table.Columns.Contains("Route Number"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Route Number"])))
                                {
                                    _newRow["Route Number"] = Convert.ToString(dr["Route Number"]);
                                }
                                else
                                {
                                    _newRow["Route Number"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Route Number"] = "";
                            }
                            if (dr.Table.Columns.Contains("Original Driver No"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Original Driver No"])))
                                {
                                    _newRow["Original Driver No"] = Convert.ToString(dr["Original Driver No"]);
                                }
                                else
                                {
                                    _newRow["Original Driver No"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Original Driver No"] = "";
                            }
                            if (dr.Table.Columns.Contains("Correct Driver Number"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Correct Driver Number"])))
                                {
                                    _newRow["Correct Driver Number"] = Convert.ToString(dr["Correct Driver Number"]);
                                }
                                else
                                {
                                    _newRow["Correct Driver Number"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Correct Driver Number"] = "";
                            }
                            if (dr.Table.Columns.Contains("Carrier Name"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Carrier Name"])))
                                {
                                    _newRow["Carrier Name"] = Convert.ToString(dr["Carrier Name"]);
                                }
                                else
                                {
                                    _newRow["Carrier Name"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Carrier Name"] = "";
                            }
                            if (dr.Table.Columns.Contains("Address"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Address"])))
                                {
                                    _newRow["Address"] = Convert.ToString(dr["Address"]);
                                }
                                else
                                {
                                    _newRow["Address"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Address"] = "";
                            }
                            if (dr.Table.Columns.Contains("City"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["City"])))
                                {
                                    _newRow["City"] = Convert.ToString(dr["City"]);
                                }
                                else
                                {
                                    _newRow["City"] = "";
                                }
                            }
                            else
                            {
                                _newRow["City"] = "";
                            }
                            if (dr.Table.Columns.Contains("State"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["State"])))
                                {
                                    _newRow["State"] = Convert.ToString(dr["State"]);
                                }
                                else
                                {
                                    _newRow["State"] = "";
                                }
                            }
                            else
                            {
                                _newRow["State"] = "";
                            }
                            if (dr.Table.Columns.Contains("Zip"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Zip"])))
                                {
                                    _newRow["Zip"] = Convert.ToString(dr["Zip"]);
                                }
                                else
                                {
                                    _newRow["Zip"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Zip"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pieces"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                                {
                                    _newRow["Pieces"] = Convert.ToString(dr["Pieces"]);
                                }
                                else
                                {
                                    _newRow["Pieces"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pieces"] = "";
                            }
                            if (dr.Table.Columns.Contains("Miles"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Miles"])))
                                {
                                    _newRow["Miles"] = Convert.ToString(dr["Miles"]);
                                }
                                else
                                {
                                    _newRow["Miles"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Miles"] = "";
                            }
                            if (dr.Table.Columns.Contains("Delivery Zip"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery Zip"])))
                                {
                                    _newRow["Delivery Zip"] = Convert.ToString(dr["Delivery Zip"]);
                                }
                                else
                                {
                                    _newRow["Delivery Zip"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Delivery Zip"] = "";
                            }
                            if (dr.Table.Columns.Contains("Zip Code Surcharge?"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Zip Code Surcharge?"])))
                                {
                                    _newRow["Zip Code Surcharge?"] = Convert.ToString(dr["Zip Code Surcharge?"]);
                                }
                                else
                                {
                                    _newRow["Zip Code Surcharge?"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Zip Code Surcharge?"] = "";
                            }
                            if (dr.Table.Columns.Contains("Store Code"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Store Code"])))
                                {
                                    _newRow["Store Code"] = Convert.ToString(dr["Store Code"]);
                                }
                                else
                                {
                                    _newRow["Store Code"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Store Code"] = "";
                            }
                            if (dr.Table.Columns.Contains("Type of Delivery"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Type of Delivery"])))
                                {
                                    _newRow["Type of Delivery"] = Convert.ToString(dr["Type of Delivery"]);
                                }
                                else
                                {
                                    _newRow["Type of Delivery"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Type of Delivery"] = "";
                            }
                            if (dr.Table.Columns.Contains("Service Type"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Service Type"])))
                                {
                                    _newRow["Service Type"] = Convert.ToString(dr["Service Type"]);
                                }
                                else
                                {
                                    _newRow["Service Type"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Service Type"] = "";
                            }
                            if (dr.Table.Columns.Contains("Bill Rate"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Bill Rate"])))
                                {
                                    _newRow["Bill Rate"] = Convert.ToString(dr["Bill Rate"]);
                                }
                                else
                                {
                                    _newRow["Bill Rate"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Bill Rate"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pieces ACC"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pieces ACC"])))
                                {
                                    _newRow["Pieces ACC"] = Convert.ToString(dr["Pieces ACC"]);
                                }
                                else
                                {
                                    _newRow["Pieces ACC"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pieces ACC"] = "";
                            }
                            if (dr.Table.Columns.Contains("FSC"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["FSC"])))
                                {
                                    _newRow["FSC"] = Convert.ToString(dr["FSC"]);
                                }
                                else
                                {
                                    _newRow["FSC"] = "";
                                }
                            }
                            else
                            {
                                _newRow["FSC"] = "";
                            }
                            if (dr.Table.Columns.Contains("Total Bill"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Total Bill"])))
                                {
                                    _newRow["Total Bill"] = Convert.ToString(dr["Total Bill"]);
                                }
                                else
                                {
                                    _newRow["Total Bill"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Total Bill"] = "";
                            }
                            if (dr.Table.Columns.Contains("Carrier Base Pay"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Carrier Base Pay"])))
                                {
                                    _newRow["Carrier Base Pay"] = Convert.ToString(dr["Carrier Base Pay"]);
                                }
                                else
                                {
                                    _newRow["Carrier Base Pay"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Carrier Base Pay"] = "";
                            }
                            if (dr.Table.Columns.Contains("Carrier ACC"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Carrier ACC"])))
                                {
                                    _newRow["Carrier ACC"] = Convert.ToString(dr["Carrier ACC"]);
                                }
                                else
                                {
                                    _newRow["Carrier ACC"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Carrier ACC"] = "";
                            }
                            if (dr.Table.Columns.Contains("Side Notes"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Side Notes"])))
                                {
                                    _newRow["Side Notes"] = Convert.ToString(dr["Side Notes"]);
                                }
                                else
                                {
                                    _newRow["Side Notes"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Side Notes"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup requested date"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup requested date"])))
                                {
                                    _newRow["Pickup requested date"] = Convert.ToString(dr["Pickup requested date"]);
                                }
                                else
                                {
                                    _newRow["Pickup requested date"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup requested date"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup will be ready by"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup will be ready by"])))
                                {
                                    _newRow["Pickup will be ready by"] = Convert.ToString(dr["Pickup will be ready by"]);
                                }
                                else
                                {
                                    _newRow["Pickup will be ready by"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup will be ready by"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup no later than"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup no later than"])))
                                {
                                    _newRow["Pickup no later than"] = Convert.ToString(dr["Pickup no later than"]);
                                }
                                else
                                {
                                    _newRow["Pickup no later than"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup no later than"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup actual date"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup actual date"])))
                                {
                                    _newRow["Pickup actual date"] = Convert.ToString(dr["Pickup actual date"]);
                                }
                                else
                                {
                                    _newRow["Pickup actual date"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup actual date"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup actual arrival time"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup actual arrival time"])))
                                {
                                    _newRow["Pickup actual arrival time"] = Convert.ToString(dr["Pickup actual arrival time"]);
                                }
                                else
                                {
                                    _newRow["Pickup actual arrival time"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup actual arrival time"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup actual depart time"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup actual depart time"])))
                                {
                                    _newRow["Pickup actual depart time"] = Convert.ToString(dr["Pickup actual depart time"]);
                                }
                                else
                                {
                                    _newRow["Pickup actual depart time"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup actual depart time"] = "";
                            }

                            if (dr.Table.Columns.Contains("Pickup name"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup name"])))
                                {
                                    _newRow["Pickup name"] = Convert.ToString(dr["Pickup name"]);
                                }
                                else
                                {
                                    _newRow["Pickup name"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup name"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup address"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup address"])))
                                {
                                    _newRow["Pickup address"] = Convert.ToString(dr["Pickup address"]);
                                }
                                else
                                {
                                    _newRow["Pickup address"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup address"] = "";
                            }

                            if (dr.Table.Columns.Contains("Pickup city"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup city"])))
                                {
                                    _newRow["Pickup city"] = Convert.ToString(dr["Pickup city"]);
                                }
                                else
                                {
                                    _newRow["Pickup city"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup city"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup state/province"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup state/province"])))
                                {
                                    _newRow["Pickup state/province"] = Convert.ToString(dr["Pickup state/province"]);
                                }
                                else
                                {
                                    _newRow["Pickup state/province"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup state/province"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup zip/postal code"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup zip/postal code"])))
                                {
                                    _newRow["Pickup zip/postal code"] = Convert.ToString(dr["Pickup zip/postal code"]);
                                }
                                else
                                {
                                    _newRow["Pickup zip/postal code"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup zip/postal code"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup text signature"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup text signature"])))
                                {
                                    _newRow["Pickup text signature"] = Convert.ToString(dr["Pickup text signature"]);
                                }
                                else
                                {
                                    _newRow["Pickup text signature"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup text signature"] = "";
                            }
                            if (dr.Table.Columns.Contains("Delivery requested date"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery requested date"])))
                                {
                                    _newRow["Delivery requested date"] = Convert.ToString(dr["Delivery requested date"]);
                                }
                                else
                                {
                                    _newRow["Delivery requested date"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Delivery requested date"] = "";
                            }
                            if (dr.Table.Columns.Contains("Deliver no earlier than"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Deliver no earlier than"])))
                                {
                                    _newRow["Deliver no earlier than"] = Convert.ToString(dr["Deliver no earlier than"]);
                                }
                                else
                                {
                                    _newRow["Deliver no earlier than"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Deliver no earlier than"] = "";
                            }
                            if (dr.Table.Columns.Contains("Deliver no later than"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Deliver no later than"])))
                                {
                                    _newRow["Deliver no later than"] = Convert.ToString(dr["Deliver no later than"]);
                                }
                                else
                                {
                                    _newRow["Deliver no later than"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Deliver no later than"] = "";
                            }
                            if (dr.Table.Columns.Contains("Delivery actual date"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery actual date"])))
                                {
                                    _newRow["Delivery actual date"] = Convert.ToString(dr["Delivery actual date"]);
                                }
                                else
                                {
                                    _newRow["Delivery actual date"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Delivery actual date"] = "";
                            }
                            if (dr.Table.Columns.Contains("Delivery actual arrive time"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery actual arrive time"])))
                                {
                                    _newRow["Delivery actual arrive time"] = Convert.ToString(dr["Delivery actual arrive time"]);
                                }
                                else
                                {
                                    _newRow["Delivery actual arrive time"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Delivery actual arrive time"] = "";
                            }
                            if (dr.Table.Columns.Contains("Delivery actual depart time"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery actual depart time"])))
                                {
                                    _newRow["Delivery actual depart time"] = Convert.ToString(dr["Delivery actual depart time"]);
                                }
                                else
                                {
                                    _newRow["Delivery actual depart time"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Delivery actual depart time"] = "";
                            }
                            if (dr.Table.Columns.Contains("Delivery text signature"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery text signature"])))
                                {
                                    _newRow["Delivery text signature"] = Convert.ToString(dr["Delivery text signature"]);
                                }
                                else
                                {
                                    _newRow["Delivery text signature"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Delivery text signature"] = "";
                            }
                            if (dr.Table.Columns.Contains("Requested by"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Requested by"])))
                                {
                                    _newRow["Requested by"] = Convert.ToString(dr["Requested by"]);
                                }
                                else
                                {
                                    _newRow["Requested by"] = "";

                                }
                            }
                            else
                            {
                                _newRow["Requested by"] = "";
                            }
                            if (dr.Table.Columns.Contains("Entered by"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Entered by"])))
                                {
                                    _newRow["Entered by"] = Convert.ToString(dr["Entered by"]);
                                }
                                else
                                {
                                    _newRow["Entered by"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Entered by"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup Delivery Transfer Flag"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup Delivery Transfer Flag"])))
                                {
                                    _newRow["Pickup Delivery Transfer Flag"] = Convert.ToString(dr["Pickup Delivery Transfer Flag"]);
                                }
                                else
                                {
                                    _newRow["Pickup Delivery Transfer Flag"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup Delivery Transfer Flag"] = "";
                            }

                            if (dr.Table.Columns.Contains("Weight"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Weight"])))
                                {
                                    _newRow["Weight"] = Convert.ToString(dr["Weight"]);
                                }
                                else
                                {
                                    _newRow["Weight"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Weight"] = "";
                            }
                            if (dr.Table.Columns.Contains("Insurance Amount"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Insurance Amount"])))
                                {
                                    _newRow["Insurance Amount"] = Convert.ToString(dr["Insurance Amount"]);
                                }
                                else
                                {
                                    _newRow["Insurance Amount"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Insurance Amount"] = "";
                            }
                            if (dr.Table.Columns.Contains("Master airway bill number"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Master airway bill number"])))
                                {
                                    _newRow["Master airway bill number"] = Convert.ToString(dr["Master airway bill number"]);
                                }
                                else
                                {
                                    _newRow["Master airway bill number"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Master airway bill number"] = "";
                            }
                            if (dr.Table.Columns.Contains("PO Number"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["PO Number"])))
                                {
                                    _newRow["PO Number"] = Convert.ToString(dr["PO Number"]);
                                }
                                else
                                {
                                    _newRow["PO Number"] = "";
                                }
                            }
                            else
                            {
                                _newRow["PO Number"] = "";
                            }
                            if (dr.Table.Columns.Contains("House airway bill number"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["House airway bill number"])))
                                {
                                    _newRow["House airway bill number"] = Convert.ToString(dr["House airway bill number"]);
                                }
                                else
                                {
                                    _newRow["House airway bill number"] = "";
                                }
                            }
                            else
                            {
                                _newRow["House airway bill number"] = "";
                            }

                            if (dr.Table.Columns.Contains("Item Number"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Item Number"])))
                                {
                                    _newRow["Item Number"] = Convert.ToString(dr["Item Number"]);
                                }
                                else
                                {
                                    _newRow["Item Number"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Item Number"] = "";
                            }
                            if (dr.Table.Columns.Contains("Item Description"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Item Description"])))
                                {
                                    _newRow["Item Description"] = Convert.ToString(dr["Item Description"]);
                                }
                                else
                                {
                                    _newRow["Item Description"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Item Description"] = "";
                            }
                            if (dr.Table.Columns.Contains("Dim Height"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Dim Height"])))
                                {
                                    _newRow["Dim Height"] = Convert.ToString(dr["Dim Height"]);
                                }
                                else
                                {
                                    _newRow["Dim Height"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Dim Height"] = "";
                            }
                            if (dr.Table.Columns.Contains("Dim Length"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Dim Length"])))
                                {
                                    _newRow["Dim Length"] = Convert.ToString(dr["Dim Length"]);
                                }
                                else
                                {
                                    _newRow["Dim Length"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Dim Length"] = "";
                            }
                            if (dr.Table.Columns.Contains("Dim Width"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Dim Width"])))
                                {
                                    _newRow["Dim Width"] = Convert.ToString(dr["Dim Width"]);
                                }
                                else
                                {
                                    _newRow["Dim Width"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Dim Width"] = "";
                            }

                            if (dr.Table.Columns.Contains("Pickup Room"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup Room"])))
                                {
                                    _newRow["Pickup Room"] = Convert.ToString(dr["Pickup Room"]);
                                }
                                else
                                {
                                    _newRow["Pickup Room"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup Room"] = "";
                            }
                            if (dr.Table.Columns.Contains("Pickup Attention"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Pickup Attention"])))
                                {
                                    _newRow["Pickup Attention"] = Convert.ToString(dr["Pickup Attention"]);
                                }
                                else
                                {
                                    _newRow["Pickup Attention"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Pickup Attention"] = "";
                            }
                            if (dr.Table.Columns.Contains("Deliver Attention"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Deliver Attention"])))
                                {
                                    _newRow["Deliver Attention"] = Convert.ToString(dr["Deliver Attention"]);
                                }
                                else
                                {
                                    _newRow["Deliver Attention"] = "";
                                }
                            }
                            else
                            {
                                _newRow["Deliver Attention"] = "";
                            }
                            dtableOrderTemplate.Rows.Add(_newRow);
                        }
                    }
                }
                catch (Exception ex)
                {
                    executionLogMessage = "OrderPostFile Creation Exception -" + ex.Message + System.Environment.NewLine;
                    executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                    objCommon.WriteErrorLog(ex, executionLogMessage);

                    ErrorResponse objErrorResponse = new ErrorResponse();
                    objErrorResponse.error = "Found exception while processing the file";
                    objErrorResponse.status = "Error";
                    objErrorResponse.code = "Excception while creating the order post file.";
                    string strErrorResponse = JsonConvert.SerializeObject(objErrorResponse);
                    DataSet dsFailureResponse = objCommon.jsonToDataSet(strErrorResponse);
                    dsFailureResponse.Tables[0].TableName = "Order-Create-Input";
                    objCommon.WriteDataToCsvFile(dsFailureResponse.Tables[0], fileName, strDatetime);
                }

                dtableOrderTemplate.TableName = "Template";
            }
        }
    }
}
