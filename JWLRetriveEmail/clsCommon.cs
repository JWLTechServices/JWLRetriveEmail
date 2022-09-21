using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Security.Cryptography;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;
using Newtonsoft.Json;

namespace JWLRetriveEmail
{
    class clsCommon
    {
        public static bool IsException = false;
        public string GetConfigValue(string Key)
        {
            string retVal = "";
            retVal = ConfigurationManager.AppSettings[Key];
            return retVal;
        }

        public struct ReturnResponse
        {
            public bool ResponseVal;
            public string Reason;

            public ReturnResponse(bool boolResponse = false)
            {
                this.ResponseVal = boolResponse;
                this.Reason = "Some Error";
            }
        }

        public struct DSResponse
        {
            public ReturnResponse dsResp;
            public DataSet DS;
        }

        public bool SendExceptionMail(string Subject, string Body)
        {
            try
            {
                string fromMail = GetConfigValue("FromMailID");
                string fromPassword = GetConfigValue("FromMailPasssword");
                string Disclaimer = GetConfigValue("MailDisclaimer");
                string toMail = GetConfigValue("ToMailID");
                return SendMail(fromMail, fromPassword, Disclaimer, toMail, "", Subject, Body, "");
            }
            catch (Exception ex)
            {
                // LogEvents(ex, "SendExceptionMail", System.Diagnostics.EventLogEntryType.Error, 190);
                WriteErrorLog(ex, "SendExceptionMail");
                return false;
            }
        }

        public bool SendMail(string fromMail, string fromPassword, string Disclaimer, string toMail, string ccMail, string Subject, string Body, string AttachmentPath)
        {
            try
            {
                string AppName = GetConfigValue("ApplicationName");
                SmtpClient smtpClient = new SmtpClient(GetConfigValue("MailSMTPHost"), Convert.ToInt32(GetConfigValue("MailSMTPPort")));
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential(fromMail, fromPassword);
                smtpClient.EnableSsl = true;

                MailAddress fromAddress = new MailAddress(fromMail);

                MailMessage mailMsg = new MailMessage();
                mailMsg.From = fromAddress;

                string[] toAddress;
                toAddress = toMail.Split(',');
                foreach (string strTo in toAddress)
                {
                    mailMsg.To.Add(strTo);
                }

                if (ccMail != "")
                {
                    string[] ccAddress;
                    ccAddress = ccMail.Split(',');
                    foreach (string strCc in ccAddress)
                    {
                        mailMsg.CC.Add(strCc);
                    }
                }

                mailMsg.Subject = Subject;

                Body = Body.Replace(System.Environment.NewLine, "<br/>");

                Body = Body + "<br/><br/>Regards,<br/>" + AppName + " <br/>Support Team<br/><br/>";

                if (Disclaimer.Trim() != "")
                {
                    Body = Body + "<br/><br/>" + Disclaimer;
                }

                mailMsg.Body = Body;
                mailMsg.IsBodyHtml = true;

                if (AttachmentPath.Trim() != "")
                {
                    Attachment att = new Attachment(AttachmentPath);
                    mailMsg.Attachments.Add(att);
                }

                smtpClient.Send(mailMsg);
                return true;
            }
            catch (Exception ex)
            {
                //LogEvents(ex, "SendMail", System.Diagnostics.EventLogEntryType.Error, 190);
                WriteErrorLog(ex, "SendMail");
                return false;
            }
        }

        public void WriteExecutionLog(string strExecutionLogMessage)
        {
            try
            {
                string AppName = GetConfigValue("ApplicationName");
                string strExecutionLogFilePath = GetConfigValue("ExecutionLogFileLocation"); ;

                if (!System.IO.Directory.Exists(strExecutionLogFilePath + @"\"))
                    System.IO.Directory.CreateDirectory(strExecutionLogFilePath + @"\");

                //    string filepath = strExecutionLogFilePath + @"\" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

                string filepath = strExecutionLogFilePath + @"\" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                //    string filepath = strExecutionLogFilePath + @"\" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                string Message = "Date/Time: " + DateTime.Now.ToString() + " " + strExecutionLogMessage + System.Environment.NewLine;


                if (!File.Exists(filepath))
                {
                    // Create a file to write to.   
                    using (StreamWriter sw = File.CreateText(filepath))
                    {
                        sw.WriteLine(Message);
                        sw.Flush();
                        sw.Close();
                    }

                }
                else
                {
                    using (StreamWriter sw = File.AppendText(filepath))
                    {
                        sw.WriteLine(Message);
                        sw.Flush();
                        sw.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog(ex, "WriteExecutionLog -" + strExecutionLogMessage);
                throw new Exception("Error in WriteExecutionLog -->" + ex.Message + ex.StackTrace);
            }
            finally
            {

            }
        }

        public void WriteErrorLog(Exception ex, string strExceptionMethod, string strExtrainfo = null)
        {

            IsException = true;
            string strErrorLogPath;

            strErrorLogPath = GetConfigValue("ErrorLogFileLocation");

            if (!System.IO.Directory.Exists(strErrorLogPath))
                System.IO.Directory.CreateDirectory(strErrorLogPath);

            string filepath = strErrorLogPath + @"\" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            string ExeMessage = "================================================" + System.Environment.NewLine;
            ExeMessage += "Date/Time: " + DateTime.Now.ToString() + System.Environment.NewLine;
            ExeMessage += "Message: " + ex.Message + System.Environment.NewLine;
            ExeMessage += "Message: " + ex.StackTrace + System.Environment.NewLine;
            if (strExceptionMethod != null)
            {
                ExeMessage += "Exception Occurred in the method: " + strExceptionMethod + System.Environment.NewLine;
            }
            if (strExtrainfo != null)
            {
                ExeMessage += System.Environment.NewLine + strExtrainfo + System.Environment.NewLine;
            }

            ExeMessage += "================================================" + System.Environment.NewLine;
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {

                    sw.WriteLine(ExeMessage);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(ExeMessage);
                }
            }

        }
        public void WriteToFile(string Message)
        {
            string path = GetConfigValue("LogfilePath");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = path + "Log_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        public void CleanAttachmentWorkingFolder()
        {
            try
            {
                string sourcePath = GetConfigValue("AttachmentWorkingFolder");

                DirectoryInfo di = new DirectoryInfo(sourcePath);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
            }
            catch (Exception ex)
            {
                string strExecutionLogMessage = "Exception in CleanAttachmentWorkingFolder" + System.Environment.NewLine;
                WriteErrorLog(ex, strExecutionLogMessage);
            }
        }

        public DSResponse GetReadEmailMappingDetails(string CustomerName, string LocationCode, string ProductCode)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramCustomerName = new SqlParameter("@CustomerName", SqlDbType.VarChar);
                paramCustomerName.Value = CustomerName;

                SqlParameter paramLocationCode = new SqlParameter("@LocationCode", SqlDbType.VarChar);
                paramLocationCode.Value = LocationCode;

                SqlParameter paramProductCode = new SqlParameter("@ProductCode", SqlDbType.VarChar);
                paramProductCode.Value = ProductCode;

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection"), CommandType.StoredProcedure, "USP_S_READEMAIL_CUSTOMERMAPPING",
                    paramCustomerName, paramLocationCode, paramProductCode);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "No Customer Details Found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetReadEmailMappingDetails");
            }
            return objResponse;
        }

        public DSResponse GetRouteStopPostTemplateDetails(string CustomerName, string LocationCode, string ProductCode)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramCustomerName = new SqlParameter("@CustomerName", SqlDbType.VarChar);
                paramCustomerName.Value = CustomerName;

                SqlParameter paramLocationCode = new SqlParameter("@LocationCode", SqlDbType.VarChar);
                paramLocationCode.Value = LocationCode;

                SqlParameter paramProductCode = new SqlParameter("@ProductCode", SqlDbType.VarChar);
                paramProductCode.Value = ProductCode;


                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection"), CommandType.StoredProcedure, "USP_S_RouteStopPostTemplate_CUSTOMERMAPPINGCOLUMNS",
                    paramCustomerName, paramLocationCode, paramProductCode);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "No Customer Details Found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetRouteStopPostTemplateDetails");
            }
            return objResponse;
        }

        public DSResponse GetRouteStopDetails(string customer_reference, string company_no)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramcustomer_reference = new SqlParameter("@customer_reference", SqlDbType.VarChar);
                paramcustomer_reference.Value = customer_reference;

                SqlParameter paramcompany_no = new SqlParameter("@company_no", SqlDbType.Int);
                paramcompany_no.Value = company_no;


                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection"), CommandType.StoredProcedure, "USP_S_RouteStopDetails",
                    paramcustomer_reference, paramcompany_no);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "Route stop details not found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetRouteStopPostTemplateDetails");
            }
            return objResponse;
        }
        public DataSet jsonToDataSet(string jsonString, string type = null)
        {
            DataSet ds = new DataSet();
            try
            {
                XmlDocument xd = new XmlDocument();
                jsonString = "{ \"rootNode\": {" + jsonString.Trim().TrimStart('{').TrimEnd('}') + "} }";
                xd = (XmlDocument)JsonConvert.DeserializeXmlNode(jsonString);

                ds.ReadXml(new XmlNodeReader(xd));

                if (type == "OrderPost")
                {

                    var UniqueId = ds.Tables[0].TableName;

                    ds.Tables[0].TableName = "id";

                    ds.Tables[0].Columns.Add("Id", typeof(System.String));

                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        //need to set value to NewColumn column
                        row["Id"] = UniqueId;   // or set it to some other value
                    }


                    for (int f = ds.Tables["Id"].ChildRelations.Count - 1; f >= 0; f--)
                    {
                        ds.Tables["Id"].ChildRelations[f].ChildTable.Constraints.Remove(ds.Tables["Id"].ChildRelations[f].RelationName);
                        ds.Tables["Id"].ChildRelations.RemoveAt(f);
                    }
                    ds.Tables["Id"].ChildRelations.Clear();
                    ds.Tables["Id"].ParentRelations.Clear();
                    ds.Tables["Id"].Constraints.Clear();


                    string columntoremove = UniqueId + "_Id";

                    if (ds.Tables.Contains("settlements"))
                    {
                        ds.Tables["settlements"].Columns.Remove(columntoremove);
                    }

                    if (ds.Tables.Contains("progress"))
                    {
                        var myTable = ds.Tables["progress"];

                        ds.Tables["progress"].Columns.Remove(ds.Tables["progress"].Columns[columntoremove]);


                        ds.Tables["progress"].Columns.Add("Id", typeof(System.String));
                        foreach (DataRow row in ds.Tables["progress"].Rows)
                        {
                            //need to set value to NewColumn column
                            row["Id"] = UniqueId;   // or set it to some other value
                        }
                    }
                }
                else if (type == "RouteStopPostAPI")
                {
                    var UniqueId = ds.Tables[0].TableName;

                    if (ds.Tables.Contains("progress"))
                    {
                        var myTable = ds.Tables["progress"];

                        //  ds.Tables["progress"].Columns.Remove(ds.Tables["progress"].Columns[columntoremove]);


                        ds.Tables["progress"].Columns.Add("id", typeof(System.String));
                        foreach (DataRow row in ds.Tables["progress"].Rows)
                        {
                            //need to set value to NewColumn column
                            row["id"] = UniqueId;   // or set it to some other value
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //WriteErrorLog(ex);
                WriteErrorLog(ex, "jsonToDataSet");
                //  LogEvents(ex, "jsonToDataSet", System.Diagnostics.EventLogEntryType.Error, 101);
                // throw new ArgumentException(ex.Message);

            }
            return ds;
        }

        public void SaveOutputDataToCsvFile<T>(List<T> reportData, string strName, string strInputFilePath, string referenceNumber, string fileName, string Datetime)
        {
            string strOutputFileLocation;
            string strOutputFile;
            try
            {

                strOutputFileLocation = GetConfigValue("OutputFilesWorkingFolder");

                if (!System.IO.Directory.Exists(strOutputFileLocation + @"\"))
                    System.IO.Directory.CreateDirectory(strOutputFileLocation + @"\");


                int fileExtPos = fileName.LastIndexOf(".");
                if (fileExtPos >= 0)
                    fileName = fileName.Substring(0, fileExtPos);


                strOutputFile = fileName + "-" + strName + "-" + Datetime;// + ".xlsx";
                strOutputFile = strOutputFileLocation + @"\" + strOutputFile + ".csv"; // ".csv";

                var lines = new List<string>();
                IEnumerable<PropertyDescriptor> props = TypeDescriptor.GetProperties(typeof(T)).OfType<PropertyDescriptor>();

                var header = string.Join(",", props.ToList().Select(x => x.Name));
                if (!File.Exists(strOutputFile))
                {
                    lines.Add(header);
                }
                var valueLines = reportData.Select(row => string.Join(",", header.Split(',').Select(a => row.GetType().GetProperty(a).GetValue(row, null))));
                lines.AddRange(valueLines);
                //File.WriteAllLines(path, lines.ToArray());

                File.AppendAllLines(strOutputFile, lines.ToArray());

            }
            catch (Exception ex)
            {
                string strExecutionLogMessage = "Exception in SaveOutputDataToCsvFile" + System.Environment.NewLine;
                WriteErrorLog(ex, strExecutionLogMessage);

            }
        }

        public void WriteDataToCsvFile(System.Data.DataTable dataTable, string fileName, string Datetime)
        {
            try
            {

                string strOutputFileLocation;
                string strOutputFile;

                strOutputFileLocation = GetConfigValue("OutputFilesWorkingFolder");

                if (!System.IO.Directory.Exists(strOutputFileLocation + @"\"))
                    System.IO.Directory.CreateDirectory(strOutputFileLocation + @"\");


                int fileExtPos = fileName.LastIndexOf(".");
                if (fileExtPos >= 0)
                    fileName = fileName.Substring(0, fileExtPos);


                strOutputFile = fileName + "-" + dataTable.TableName + "-" + Datetime;// + ".xlsx";
                strOutputFile = strOutputFileLocation + @"\" + strOutputFile + ".csv"; // ".csv";

                StringBuilder fileContent = new StringBuilder();
                StringBuilder HeaderContent = new StringBuilder();

                if (!File.Exists(strOutputFile))
                {
                    foreach (var col in dataTable.Columns)
                    {
                        HeaderContent.Append(col.ToString() + ",");
                    }
                    HeaderContent.Replace(",", System.Environment.NewLine, HeaderContent.Length - 1, 1);
                    File.WriteAllText(strOutputFile, HeaderContent.ToString());
                }

                foreach (DataRow dr in dataTable.Rows)
                {
                    foreach (var column in dr.ItemArray)
                    {
                        fileContent.Append("\"" + column.ToString() + "\",");
                    }

                    fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);

                }
                File.AppendAllText(strOutputFile, fileContent.ToString());
            }
            catch (Exception ex)
            {
                string strExecutionLogMessage = "Exception in WriteDataToCsvFile" + System.Environment.NewLine;
                WriteErrorLog(ex, strExecutionLogMessage);

            }
        }

        public DSResponse GetOrderPostTemplateDetails(string CustomerName, string LocationCode, string ProductCode, string ProductSubCode)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramCustomerName = new SqlParameter("@CustomerName", SqlDbType.VarChar);
                paramCustomerName.Value = CustomerName;

                SqlParameter paramLocationCode = new SqlParameter("@LocationCode", SqlDbType.VarChar);
                paramLocationCode.Value = LocationCode;

                SqlParameter paramProductCode = new SqlParameter("@ProductCode", SqlDbType.VarChar);
                paramProductCode.Value = ProductCode;

                SqlParameter paramProductSubCode = new SqlParameter("@ProductSubCode", SqlDbType.VarChar);
                if (string.IsNullOrEmpty(ProductSubCode))
                {
                    paramProductSubCode.Value = DBNull.Value;
                }
                else
                {
                    paramProductSubCode.Value = ProductSubCode;
                }

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection"), CommandType.StoredProcedure, "USP_S_OrderPostTemplate_CustomerMappingColumns",
                    paramCustomerName, paramLocationCode, paramProductCode, paramProductSubCode);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "No Customer Details Found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetOrderPostTemplateDetails");
            }
            return objResponse;
        }

        public DSResponse GetOrderDetails(string customer_reference, string company_no)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramcustomer_reference = new SqlParameter("@customer_reference", SqlDbType.VarChar);
                paramcustomer_reference.Value = customer_reference;

                SqlParameter paramcompany_no = new SqlParameter("@company_no", SqlDbType.Int);
                paramcompany_no.Value = company_no;


                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection"), CommandType.StoredProcedure, "USP_S_OrderDetails",
                    paramcustomer_reference, paramcompany_no);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "Order details not found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetOrderDetails");
            }
            return objResponse;
        }

        public DSResponse GetTGTCustomerMappingDetails(string strtype, string strOriginAddress, string strDestinationAddress)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramtype = new SqlParameter("@Type", SqlDbType.VarChar);
                paramtype.Value = strtype;

                SqlParameter paramOriginAddress = new SqlParameter("@OriginAddress", SqlDbType.VarChar);
                paramOriginAddress.Value = strOriginAddress;

                SqlParameter paramDestinationAddress = new SqlParameter("@DestinationAddress", SqlDbType.VarChar);
                paramDestinationAddress.Value = strDestinationAddress;

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection"), CommandType.StoredProcedure, "USP_S_CustomerMappingDetails",
                    paramtype, paramOriginAddress, paramDestinationAddress);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "Customer NBR Mapping details not found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetTGTCustomerMappingDetails");
            }
            return objResponse;
        }


        public DSResponse GetBillingRatesAndPayableRates_CustomerMappingDetails(string company, string customerNumber)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramCompany = new SqlParameter("@Company", SqlDbType.Int);
                paramCompany.Value = company;

                SqlParameter paramCustomerNumber = new SqlParameter("@CustomerNumber", SqlDbType.Int);
                paramCustomerNumber.Value = customerNumber;

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection"), CommandType.StoredProcedure, "USP_S_BillingRates_PayableRates_CustomerMapping",
                    paramCompany, paramCustomerNumber);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "Billing Rates And Payable Rates customer mapping details not found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetBillingRatesAndPayableRates_CustomerMappingDetails");
            }
            return objResponse;
        }
    }
}
