using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EAGetMail;
using Newtonsoft.Json;

namespace JWLRetriveEmail
{
    class clsODCDS : clsCommon
    {
        public ReturnResponse ProcessCDSODData(EAGetMail.Mail oMail, EAGetMail.Attachment att, string customerName, string locationCode, string productCode, string productSubCode,
            string fileName, string extn, string datetime)
        {
            ReturnResponse returnResponse = new ReturnResponse();
            string executionLogMessage = "";
            string emailSubject = "";
            clsCommon objCommon = new clsCommon();
            try
            {
                clsCommon.DSResponse objCMDsResponse = new clsCommon.DSResponse();
                objCMDsResponse = objCommon.GetReadEmailMappingDetails(customerName, locationCode, productCode);
                if (objCMDsResponse.dsResp.ResponseVal)
                {
                    string tempfilePath = objCommon.GetConfigValue("AttachmentWorkingFolder");
                    string attachmentPath = objCommon.GetConfigValue("AutomationFileLocation");
                    string filePath = Convert.ToString(objCMDsResponse.DS.Tables[0].Rows[0]["FileLocation"]);
                    string company_no = Convert.ToString(objCMDsResponse.DS.Tables[0].Rows[0]["CompanyNumber"]);
                    string customerNumber = "";

                    DataTable dtCDSBadData = new DataTable();
                    DataTable dtCDSMissingConfigData = new DataTable();
                    string customerrefmappingcolumnname = string.Empty;

                    attachmentPath = attachmentPath + @"\" + locationCode + @"\" + filePath;

                    if (!System.IO.Directory.Exists(tempfilePath))
                        System.IO.Directory.CreateDirectory(tempfilePath);


                    string attname = String.Format("{0}\\{1}", tempfilePath, att.Name);
                    att.SaveAs(attname, true);

                    DataSet dsExcel = new DataSet();
                    if (extn.ToLower().Contains(".csv"))
                    {
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

                    DataTable dtConfiguredData = new DataTable();
                    clsCommon.DSResponse objDsRes = new clsCommon.DSResponse();
                    objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);
                    if (objDsRes.dsResp.ResponseVal)
                    {
                        dtConfiguredData = objDsRes.DS.Tables[0];
                        customerNumber = Convert.ToString(dtConfiguredData.Rows[0]["CustomerNumber"]);
                        customerrefmappingcolumnname = Convert.ToString(dtConfiguredData.Rows[0]["Customer_Reference"]);

                        dtCDSBadData = dsExcel.Tables[0].Copy();
                        dtCDSMissingConfigData = dsExcel.Tables[0].Copy();
                        dsExcel = CDSGenerateOrderDataTable(dsExcel.Tables[0], dtConfiguredData, productSubCode);

                        if (dsExcel.Tables.Count > 0)
                        {
                            DataTable dataTable = dsExcel.Tables[0];

                            if (dataTable.Rows.Count < 1)
                            {
                                executionLogMessage = "No data found in the file  :  " + attname + System.Environment.NewLine;
                                executionLogMessage += "But Still processed this file";
                                objCommon.WriteExecutionLog(executionLogMessage);

                            }

                            // clsCommon.DSResponse objDsResponse = new clsCommon.DSResponse();

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


                            dtableOrderTemplate.Columns.Add("add_charge_amt1");
                            dtableOrderTemplate.Columns.Add("add_charge_code1");
                            dtableOrderTemplate.Columns.Add("add_charge_amt2");
                            dtableOrderTemplate.Columns.Add("add_charge_code2");
                            dtableOrderTemplate.Columns.Add("add_charge_amt3");
                            dtableOrderTemplate.Columns.Add("add_charge_code3");
                            dtableOrderTemplate.Columns.Add("add_charge_amt4");
                            dtableOrderTemplate.Columns.Add("add_charge_code4");

                            dtableOrderTemplate.Columns.Add("add_charge_amt5");
                            dtableOrderTemplate.Columns.Add("add_charge_code5");
                            dtableOrderTemplate.Columns.Add("add_charge_amt6");
                            dtableOrderTemplate.Columns.Add("add_charge_code6");
                            dtableOrderTemplate.Columns.Add("Pickup special instr long");
                            dtableOrderTemplate.Columns.Add("status_code");
                            dtableOrderTemplate.Columns.Add("settlement_pct");

                            dtableOrderTemplate.Columns.Add("UnitCount");
                            dtableOrderTemplate.Columns.Add("Actual Receive Date");
                            dtableOrderTemplate.Columns.Add("Stairs");
                            dtableOrderTemplate.Columns.Add("Manufacturer");
                            dtableOrderTemplate.Columns.Add("Delivery Type");
                            dtableOrderTemplate.Columns.Add("Men");


                            // clsCommon.DSResponse objDsRes = new clsCommon.DSResponse();
                            objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);
                            if (objDsRes.dsResp.ResponseVal)
                            {
                                // string strDatetime = DateTime.Now.ToString("yyyyMMddHHmmss");
                                DataTable dtOrderData = dsExcel.Tables[0];

                                clsCommon.ReturnResponse objreturnResponseGenOrdTemplate = new clsCommon.ReturnResponse();
                                objreturnResponseGenOrdTemplate = GenerateOrderTemplate(ref dtableOrderTemplate, dtOrderData, fileName, customerName, locationCode, productCode, productSubCode, datetime);
                                if (objreturnResponseGenOrdTemplate.ResponseVal)
                                {
                                    //   DataTable dtableOrderTemplate1 = new DataTable();
                                    DataTable dtableOrderTemplateFinal = new DataTable();

                                    dtableOrderTemplateFinal = dtableOrderTemplate.Copy();
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
                                        if (objCommon.GetConfigValue("CDSEnableNewCalculation") == "Y")
                                        {
                                            GenerateCDSOrderTemplate(ref dtableOrderTemplateFinal, dtBillingRates, dtPayableRates, dtCDSBadData, dtCDSMissingConfigData, company_no, customerNumber, fileName, datetime, customerrefmappingcolumnname);
                                        }
                                        else
                                        {
                                            DataTable dtFSCRates = new DataTable();
                                            DataTable dtFSCRatesfromDB = new DataTable();
                                            DataTable tblFSCRatesFiltered = new DataTable();
                                            string fscRateDetailsfilepath = objCommon.GetConfigValue("FSCRatesCustomerMappingFilepath");
                                            DataSet dsFscData = clsExcelHelper.ImportExcelXLSXToDataSet_FSCRATES(fscRateDetailsfilepath, true, Convert.ToInt32(company_no), Convert.ToInt32(customerNumber));
                                            if (dsFscData != null && dsFscData.Tables[0].Rows.Count > 0)
                                            {
                                                dtFSCRates = dsFscData.Tables["FSCRatesMapping$"];
                                            }
                                            else
                                            {
                                                executionLogMessage = "DT-Diesel price data not found in this file " + fscRateDetailsfilepath + System.Environment.NewLine;
                                                executionLogMessage += "So not able to process this file, please update the fsc sheet with appropriate values";
                                                executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                objCommon.WriteExecutionLog(executionLogMessage);

                                                string fromMail = objCommon.GetConfigValue("FromMailID");
                                                string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                                string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                                string toMail = objCommon.GetConfigValue("CDSSendMissingDieselPriceEmailTo");
                                                string subject = "DT-Diesel price data not found in this file " + fscRateDetailsfilepath;
                                                objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                                                throw new NullReferenceException("Diesel price data not found in this file " + fscRateDetailsfilepath);
                                            }

                                            //DataTable dtCarrierFSCBillableRates = new DataTable();
                                            DataTable dtCarrierFSCRatesfromDB = new DataTable();
                                            DataTable tblCarrierFSCRatesFiltered = new DataTable();


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


                                                int miles = 0;
                                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Miles"])))
                                                {
                                                    miles = Convert.ToInt32(dr["Miles"]);
                                                    if (miles < 0)
                                                    {
                                                        dr["Miles"] = 0;
                                                        miles = 0;
                                                    }
                                                }

                                                DataTable tblBillRatesFiltered = new DataTable();

                                                IEnumerable<DataRow> billratesfilteredRows = dtBillingRates.AsEnumerable()
                                                .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                                                         && row.Field<string>("DeliveryName") == deliveryName
                                                          && (row.Field<decimal>("minimum_miles") <= miles)
                                                          && (miles <= row.Field<decimal>("maximum_miles")));

                                                if (billratesfilteredRows.Any())
                                                {
                                                    tblBillRatesFiltered = billratesfilteredRows.CopyToDataTable();
                                                }
                                                else
                                                {
                                                    billratesfilteredRows = dtBillingRates.AsEnumerable()
                                             .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                                             && row.Field<string>("DeliveryName") is null
                                              && (row.Field<decimal>("minimum_miles") <= miles)
                                              && (miles <= row.Field<decimal>("maximum_miles")));

                                                    if (billratesfilteredRows.Any())
                                                    {
                                                        tblBillRatesFiltered = billratesfilteredRows.CopyToDataTable();
                                                    }
                                                }

                                                DataTable tblPayableRatesFiltered = new DataTable();
                                                IEnumerable<DataRow> payableratesfilteredRows = dtPayableRates.AsEnumerable()
                                                .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                                                && row.Field<string>("DeliveryName") == deliveryName
                                                && (row.Field<decimal>("minimum_miles") <= miles)
                                                && (miles <= row.Field<decimal>("maximum_miles")));


                                                if (payableratesfilteredRows.Any())
                                                {
                                                    tblPayableRatesFiltered = payableratesfilteredRows.CopyToDataTable();
                                                }
                                                else
                                                {
                                                    payableratesfilteredRows = dtPayableRates.AsEnumerable()
                                                .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                                                && row.Field<string>("DeliveryName") is null
                                                && (row.Field<decimal>("minimum_miles") <= miles)
                                                && (miles <= row.Field<decimal>("maximum_miles")));

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
                                                        executionLogMessage = "DT-Diesel price is missing for date  " + dtdeliveryDate.ToShortDateString() + System.Environment.NewLine;
                                                        executionLogMessage += "So not able to process this file, please update the fsc sheet with appropriate values." + System.Environment.NewLine;
                                                        executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                        objCommon.WriteExecutionLog(executionLogMessage);
                                                        string fromMail = objCommon.GetConfigValue("FromMailID");
                                                        string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                                        string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                                        string toMail = objCommon.GetConfigValue("CDSSendMissingDieselPriceEmailTo");
                                                        string subject = "DT-Diesel price is missing for date  " + dtdeliveryDate.ToShortDateString();
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
                                                                executionLogMessage = "DT-FSC Billing Rates missing for this fuel charge   " + fuelcharge + System.Environment.NewLine;
                                                                executionLogMessage += "So not able to process this file, please update the billable rates mapping table with appropriate values";
                                                                executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                                objCommon.WriteExecutionLog(executionLogMessage);
                                                                string fromMail = objCommon.GetConfigValue("FromMailID");
                                                                string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                                                string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                                                string toMail = objCommon.GetConfigValue("ToMailID");
                                                                string subject = "DT-FSC Billing Rates missing for this fuel charge   " + fuelcharge;
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
                                                                executionLogMessage = "DT-FSC Payable Rates missing for this fuel charge   " + fuelcharge + System.Environment.NewLine;
                                                                executionLogMessage += "So not able to process this file, please update the payable rates mapping table with appropriate values";
                                                                executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                                                objCommon.WriteExecutionLog(executionLogMessage);
                                                                string fromMail = objCommon.GetConfigValue("FromMailID");
                                                                string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                                                string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                                                string toMail = objCommon.GetConfigValue("CDSSendMissingConfEmailTo");
                                                                string subject = "DT-FSC Billing Rates missing for this fuel charge   " + fuelcharge;
                                                                objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                                                                throw new NullReferenceException("Diesel prices missing for date  " + dtdeliveryDate);
                                                            }
                                                        }
                                                    }
                                                }

                                                if (tblBillRatesFiltered.Rows.Count > 0)
                                                {
                                                    double billRate = 0;
                                                    double minimumRate = 0; ;

                                                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["minimum_rate"])))
                                                    {
                                                        minimumRate = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["minimum_rate"]); ;
                                                    }

                                                    if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                                                    {
                                                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"])))
                                                        {
                                                            billRate = Convert.ToDouble(1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]));
                                                            if (billRate < minimumRate)
                                                            {
                                                                billRate = minimumRate;
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
                                                            if (billRate < minimumRate)
                                                            {
                                                                billRate = minimumRate;
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
                                                    double minimumRate = 0;
                                                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["minimum_rate"])))
                                                    {
                                                        minimumRate = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["minimum_rate"]);
                                                    }

                                                    if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                                                    {
                                                        if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                                                        {
                                                            carrierBasePay = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);
                                                            if (carrierBasePay < minimumRate)
                                                            {
                                                                carrierBasePay = minimumRate;
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
                                                            if (carrierBasePay < minimumRate)
                                                            {
                                                                carrierBasePay = minimumRate;
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
                                    }

                                    dtableOrderTemplateFinal.TableName = "Template";
                                    //clsExcelHelper.ExportDataToXLSX(dtableOrderTemplateFinal, attachmentPath, fileName);
                                    clsExcelHelperNew.ExportDataTableToXLSX(dtableOrderTemplateFinal, attachmentPath, fileName);

                                    returnResponse.ResponseVal = true;
                                    objCommon.CleanAttachmentWorkingFolder();
                                }
                                else
                                {
                                    executionLogMessage = "DT-Issue in Generating Order Template " + System.Environment.NewLine;
                                    executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                    executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                    executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                    executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                    emailSubject = "DT-Issue in Generating Order Template";
                                    objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                    objCommon.WriteExecutionLog(executionLogMessage);
                                    returnResponse.Reason = emailSubject;
                                    returnResponse.ResponseVal = false;
                                }
                            }
                            else
                            {
                                executionLogMessage = "DT-Order Post Template Mapping Missing " + System.Environment.NewLine;
                                executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                                executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                                executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                                executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                emailSubject = "DT-Order Post Template Mapping Missing";
                                objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                                objCommon.WriteExecutionLog(executionLogMessage);
                                returnResponse.Reason = emailSubject;
                                returnResponse.ResponseVal = false;
                            }
                        }
                        else
                        {
                            executionLogMessage = "DT-No data found after Export, Please check the file " + System.Environment.NewLine;
                            executionLogMessage += "Attachment Name -" + fileName + System.Environment.NewLine;
                            executionLogMessage = "From :  " + oMail.From.ToString() + System.Environment.NewLine;
                            executionLogMessage = "Email Address :  " + oMail.From.Address.ToString() + System.Environment.NewLine;
                            executionLogMessage += "Subject : " + oMail.Subject + System.Environment.NewLine;
                            executionLogMessage += "ReceivedDate : " + oMail.ReceivedDate + System.Environment.NewLine;
                            emailSubject = "DT-No data found for this Excel Post Convertion to data set";
                            objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                            objCommon.WriteExecutionLog(executionLogMessage);
                            returnResponse.Reason = emailSubject;
                            returnResponse.ResponseVal = false;
                            // continue;
                        }
                    }
                    else
                    {
                        executionLogMessage = "DT-Order Post Template Mapping Missing " + System.Environment.NewLine;
                        executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                        executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                        executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                        executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                        emailSubject = "DT-Order Post Template Mapping Missing";
                        objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                        objCommon.WriteExecutionLog(executionLogMessage);
                        returnResponse.Reason = emailSubject;
                        returnResponse.ResponseVal = false;
                        // continue;
                    }

                }
                else
                {
                    executionLogMessage = "DT-Email Mapping Missing For this Customer" + System.Environment.NewLine;
                    executionLogMessage += "CustomerName -" + customerName + System.Environment.NewLine;
                    executionLogMessage += "LocationCode -" + locationCode + System.Environment.NewLine;
                    executionLogMessage += "ProductCode -" + productCode + System.Environment.NewLine;
                    executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                    emailSubject = "DT-Email Mapping Missing For " + customerName;
                    objCommon.SendExceptionMail(emailSubject, executionLogMessage);
                    objCommon.WriteExecutionLog(executionLogMessage);
                    returnResponse.Reason = emailSubject;
                    returnResponse.ResponseVal = false;
                    //continue;
                }

            }
            catch (Exception ex)
            {
                executionLogMessage = "DT-Exception in ProcessCDSODData" + System.Environment.NewLine;
                WriteErrorLog(ex, executionLogMessage);
                returnResponse.Reason = ex.Message;
                returnResponse.ResponseVal = false;
            }
            return returnResponse;
        }

        private static DataSet CDSGenerateOrderDataTable(DataTable dtinputData, DataTable dtConfiguredData, string type)
        {


            dtinputData = dtinputData.AsEnumerable()
                .OrderBy(en => en.Field<string>("Actual Delivery Date"))
                .ThenBy(en => en.Field<string>("CDS#")).CopyToDataTable();


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

            dtinputData.Columns.Add("Delivery_text_signature_Value");
            dtinputData.Columns.Add("Pickup_address_Value");
            dtinputData.Columns.Add("Pickup_city_Value");
            dtinputData.Columns.Add("Pickup_state/province_Value");
            dtinputData.Columns.Add("Pickup_zip/postal_code_Value");
            dtinputData.Columns.Add("Pickup_text_signature_Value");


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

        private static clsCommon.ReturnResponse GenerateOrderTemplate(ref DataTable dtableOrderTemplate, DataTable dtOrderData,
            string fileName, string customerName, string locationCode,
            string productCode, string productSubCode, string datetime)
        {

            clsCommon objCommon = new clsCommon();


            clsCommon.ReturnResponse returnResponse = new clsCommon.ReturnResponse();
            string executionLogMessage = string.Empty;
            //  DataTable dtableOrderTemplate = new DataTable();
            dtableOrderTemplate.Clear();

            clsCommon.DSResponse objDsRes = new clsCommon.DSResponse();
            objDsRes = objCommon.GetOrderPostTemplateDetails(customerName, locationCode, productCode, productSubCode);

            if (objDsRes.dsResp.ResponseVal)
            {
                //  string strDatetime = DateTime.Now.ToString("yyyyMMddHHmmss");
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

                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_amt1"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_amt1"])].ColumnName = "add_charge_amt1";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code1"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code1"])].ColumnName = "add_charge_code1";
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_amt2"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_amt2"])].ColumnName = "add_charge_amt2";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code2"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code2"])].ColumnName = "add_charge_code2";
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_amt3"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_amt3"])].ColumnName = "add_charge_amt3";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code3"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code3"])].ColumnName = "add_charge_code3";
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_amt4"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_amt4"])].ColumnName = "add_charge_amt4";
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code4"])))
                    {
                        dtOrderData.Columns[Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code4"])].ColumnName = "add_charge_code4";
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

                            if (dr.Table.Columns.Contains("add_charge_amt1"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_amt1"])))
                                {
                                    _newRow["add_charge_amt1"] = Convert.ToString(dr["add_charge_amt1"]);
                                }
                                else
                                {
                                    _newRow["add_charge_amt1"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_amt1"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_code1"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_code1"])))
                                {
                                    _newRow["add_charge_code1"] = Convert.ToString(dr["add_charge_code1"]);
                                }
                                else
                                {
                                    _newRow["add_charge_code1"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_code1"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_amt2"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_amt2"])))
                                {
                                    _newRow["add_charge_amt2"] = Convert.ToString(dr["add_charge_amt2"]);
                                }
                                else
                                {
                                    _newRow["add_charge_amt2"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_amt2"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_code2"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_code2"])))
                                {
                                    _newRow["add_charge_code2"] = Convert.ToString(dr["add_charge_code2"]);
                                }
                                else
                                {
                                    _newRow["add_charge_code2"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_code2"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_amt3"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_amt3"])))
                                {
                                    _newRow["add_charge_amt3"] = Convert.ToString(dr["add_charge_amt3"]);
                                }
                                else
                                {
                                    _newRow["add_charge_amt3"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_amt3"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_code3"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_code3"])))
                                {
                                    _newRow["add_charge_code3"] = Convert.ToString(dr["add_charge_code3"]);
                                }
                                else
                                {
                                    _newRow["add_charge_code3"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_code3"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_amt4"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_amt4"])))
                                {
                                    _newRow["add_charge_amt4"] = Convert.ToString(dr["add_charge_amt4"]);
                                }
                                else
                                {
                                    _newRow["add_charge_amt4"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_amt4"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_code4"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_code4"])))
                                {
                                    _newRow["add_charge_code4"] = Convert.ToString(dr["add_charge_code4"]);
                                }
                                else
                                {
                                    _newRow["add_charge_code4"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_code4"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_amt5"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_amt5"])))
                                {
                                    _newRow["add_charge_amt5"] = Convert.ToString(dr["add_charge_amt5"]);
                                }
                                else
                                {
                                    _newRow["add_charge_amt5"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_amt5"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_code5"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_code5"])))
                                {
                                    _newRow["add_charge_code5"] = Convert.ToString(dr["add_charge_code5"]);
                                }
                                else
                                {
                                    _newRow["add_charge_code5"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_code5"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_amt6"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_amt6"])))
                                {
                                    _newRow["add_charge_amt6"] = Convert.ToString(dr["add_charge_amt6"]);
                                }
                                else
                                {
                                    _newRow["add_charge_amt6"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_amt6"] = "";
                            }

                            if (dr.Table.Columns.Contains("add_charge_code6"))
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["add_charge_code6"])))
                                {
                                    _newRow["add_charge_code6"] = Convert.ToString(dr["add_charge_code6"]);
                                }
                                else
                                {
                                    _newRow["add_charge_code6"] = "";
                                }
                            }
                            else
                            {
                                _newRow["add_charge_code6"] = "";
                            }

                            if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code1_Value"])))
                            {
                                _newRow["add_charge_code1"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code1_Value"]);
                            }
                            if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code2_Value"])))
                            {
                                _newRow["add_charge_code2"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code2_Value"]);
                            }
                            if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code3_Value"])))
                            {
                                _newRow["add_charge_code3"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code3_Value"]);
                            }
                            if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code4_Value"])))
                            {
                                _newRow["add_charge_code4"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code4_Value"]);
                            }
                            if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code5_Value"])))
                            {
                                _newRow["add_charge_code5"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code5_Value"]);
                            }
                            if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code6_Value"])))
                            {
                                _newRow["add_charge_code6"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["add_charge_code6_Value"]);
                            }

                            if (!string.IsNullOrEmpty(Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["settlement_pct_Value"])))
                            {
                                _newRow["settlement_pct"] = Convert.ToString(objDsRes.DS.Tables[0].Rows[0]["settlement_pct_Value"]);
                            }

                            if (customerName == objCommon.GetConfigValue("CDSCustomerName"))
                            {
                                if (dr.Table.Columns.Contains("UnitCount"))
                                {
                                    if (!string.IsNullOrEmpty(Convert.ToString(dr["UnitCount"])))
                                    {
                                        _newRow["UnitCount"] = Convert.ToString(dr["UnitCount"]);
                                    }
                                    else
                                    {
                                        _newRow["UnitCount"] = "";
                                    }
                                }
                                else
                                {
                                    _newRow["UnitCount"] = "";
                                }

                                if (dr.Table.Columns.Contains("Actual Receive Date"))
                                {
                                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Actual Receive Date"])))
                                    {
                                        _newRow["Actual Receive Date"] = Convert.ToString(dr["Actual Receive Date"]);
                                    }
                                    else
                                    {
                                        _newRow["Actual Receive Date"] = "";
                                    }
                                }
                                else
                                {
                                    _newRow["Actual Receive Date"] = "";
                                }

                                if (dr.Table.Columns.Contains("Stairs"))
                                {
                                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Stairs"])))
                                    {
                                        _newRow["Stairs"] = Convert.ToString(dr["Stairs"]);
                                    }
                                    else
                                    {
                                        _newRow["Stairs"] = "";
                                    }
                                }
                                else
                                {
                                    _newRow["Stairs"] = "";
                                }

                                if (dr.Table.Columns.Contains("Manufacturer"))
                                {
                                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Manufacturer"])))
                                    {
                                        _newRow["Manufacturer"] = Convert.ToString(dr["Manufacturer"]);
                                    }
                                    else
                                    {
                                        _newRow["Manufacturer"] = "";
                                    }
                                }
                                else
                                {
                                    _newRow["Manufacturer"] = "";
                                }

                                if (dr.Table.Columns.Contains("Delivery Type"))
                                {
                                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery Type"])))
                                    {
                                        _newRow["Delivery Type"] = Convert.ToString(dr["Delivery Type"]);
                                    }
                                    else
                                    {
                                        _newRow["Delivery Type"] = "";
                                    }
                                }
                                else
                                {
                                    _newRow["Delivery Type"] = "";
                                }


                                if (dr.Table.Columns.Contains("Men"))
                                {
                                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Men"])))
                                    {
                                        _newRow["Men"] = Convert.ToString(dr["Men"]);
                                    }
                                    else
                                    {
                                        _newRow["Men"] = "";
                                    }
                                }
                                else
                                {
                                    _newRow["Men"] = "";
                                }
                            }

                            dtableOrderTemplate.Rows.Add(_newRow);
                        }
                    }
                    returnResponse.ResponseVal = true;
                }
                catch (Exception ex)
                {
                    returnResponse.ResponseVal = false;
                    executionLogMessage = "DT-OrderPostFile Creation Exception -" + ex.Message + System.Environment.NewLine;
                    executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                    objCommon.WriteErrorLog(ex, executionLogMessage);

                    ErrorResponse objErrorResponse = new ErrorResponse();
                    objErrorResponse.error = "DT-Found exception while processing the file";
                    objErrorResponse.status = "Error";
                    objErrorResponse.code = "Excception while creating the order post file.";
                    string strErrorResponse = JsonConvert.SerializeObject(objErrorResponse);
                    DataSet dsFailureResponse = objCommon.jsonToDataSet(strErrorResponse);
                    dsFailureResponse.Tables[0].TableName = "Order-Create-Input";
                    objCommon.WriteDataToCsvFile(dsFailureResponse.Tables[0], fileName, datetime);
                }

                dtableOrderTemplate.TableName = "Template";
            }
            return returnResponse;
        }

        private static clsCommon.ReturnResponse GenerateCDSOrderTemplate(ref DataTable dtableOrderTemplateFinal, DataTable dtBillingRates, DataTable dtPayableRates,
         DataTable dtCDSBadData, DataTable dtCDSMissingConfigData, string company_no, string customerNumber, string fileName, string dateTime, string customerrefmappingcolumnname)
        {

            clsCommon objCommon = new clsCommon();

            clsCommon.ReturnResponse returnResponse = new clsCommon.ReturnResponse();
            string executionLogMessage = string.Empty;
            DataTable dtFSCRates = new DataTable();
            DataTable dtFSCRatesfromDB = new DataTable();
            DataTable tblFSCRatesFiltered = new DataTable();
            string fscRateDetailsfilepath = objCommon.GetConfigValue("FSCRatesCustomerMappingFilepath");
            DataSet dsFscData = clsExcelHelper.ImportExcelXLSXToDataSet_FSCRATES(fscRateDetailsfilepath, true, Convert.ToInt32(company_no), Convert.ToInt32(customerNumber));
            if (dsFscData != null && dsFscData.Tables[0].Rows.Count > 0)
            {
                dtFSCRates = dsFscData.Tables["FSCRatesMapping$"];
            }
            else
            {
                executionLogMessage = "DT-Diesel price data not found in this file " + fscRateDetailsfilepath + System.Environment.NewLine;
                executionLogMessage += "So not able to process this file, please update the fsc sheet with appropriate values";
                executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                objCommon.WriteExecutionLog(executionLogMessage);

                string fromMail = objCommon.GetConfigValue("FromMailID");
                string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                string toMail = objCommon.GetConfigValue("CDSSendMissingDieselPriceEmailTo");
                string subject = "DT-Diesel price data not found in this file " + fscRateDetailsfilepath;
                objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                throw new NullReferenceException("Diesel price data not found in this file " + fscRateDetailsfilepath);
            }

            DataTable dtCarrierFSCRatesfromDB = new DataTable();
            DataTable tblCarrierFSCRatesFiltered = new DataTable();



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

            var itemsToDelete = new List<DataRow>();

            DataTable dtBadData = new DataTable();
            dtBadData.Clear();
            dtBadData.Columns.Add("Customer Reference");

            DataTable dtMissingConfData = new DataTable();
            dtMissingConfData.Clear();
            dtMissingConfData.Columns.Add("Customer Reference");

            foreach (DataRow dr in dtableOrderTemplateFinal.Rows)
            {
                string deliveryType = string.Empty;
                if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery Type"])))
                {
                    deliveryType = Convert.ToString(dr["Delivery Type"]);
                }

                object delivertyDate = dr["Delivery Date"];
                object receivedDate;

                if (deliveryType != null && deliveryType.ToUpper() == "RETURN")
                    receivedDate = dr["Delivery Date"];
                else
                    receivedDate = dr["Actual Receive Date"];


                if (delivertyDate == DBNull.Value || receivedDate == DBNull.Value)
                    break;

                DateTime dtdeliveryDate = Convert.ToDateTime(Regex.Replace(delivertyDate.ToString(), @"\t", ""));
                DateTime dtreceivedDate = Convert.ToDateTime(Regex.Replace(receivedDate.ToString(), @"\t", ""));


                //    int diffDays = Convert.ToInt32((dtreceivedDate - dtdeliveryDate).TotalDays);
                int diffDays = Convert.ToInt32((dtdeliveryDate - dtreceivedDate).TotalDays);
                //if (deliveryType != null && deliveryType.ToUpper() == "RETURN")
                //{
                //    diffDays = 0;
                //}

                var invCulture = System.Globalization.CultureInfo.InvariantCulture;
                //string deliveryName = Convert.ToString(dr["Customer Name"]);
                string deliveryName = Convert.ToString(dr["Manufacturer"]);


                //if ((!string.IsNullOrEmpty(Convert.ToString(dr["Pieces"]))) && (!string.IsNullOrEmpty(Convert.ToString(dr["UnitCount"]))))
                //{
                //    if (Convert.ToInt32(dr["Pieces"]) == 0 && Convert.ToInt32(dr["UnitCount"]) > 0)
                //    {
                //        deliveryName = "Millwork";
                //    }
                //}

                if (string.IsNullOrEmpty(Convert.ToString(dr["Pieces"])))
                {
                    dr["Pieces"] = 0;
                }
                if (string.IsNullOrEmpty(Convert.ToString(dr["UnitCount"])))
                {
                    dr["UnitCount"] = 0;
                }


                if (Convert.ToInt32(dr["Pieces"]) > 0 && Convert.ToInt32(dr["UnitCount"]) > 0)
                {
                    // write the details into event and continue

                    // DataTable dtBadData = new DataTable();
                    // dtBadData = dtableOrderTemplateFinal.Select("[Customer Reference]= '" + dr["Customer Reference"] + "'").CopyToDataTable();
                    DataRow _newRow = dtBadData.NewRow();
                    _newRow["Customer Reference"] = dr["Customer Reference"];
                    dtBadData.Rows.Add(_newRow);

                    //dtBadData.TableName = "BadData";
                    //string badDataFilePath = objCommon.GetConfigValue("BadDataFileFolder");
                    //objCommon.WriteDataToCsvFile(dtBadData, badDataFilePath, fileName, dateTime);

                    executionLogMessage = "DT-CDS Order Template File Generation Error " + System.Environment.NewLine;
                    executionLogMessage += "Cabinet Count and Unit Count found greater than 0 for this record" + System.Environment.NewLine; ;
                    executionLogMessage += "Customer Reference -" + dr["Customer Reference"] + System.Environment.NewLine;
                    executionLogMessage += "Cabinet Count -" + dr["Pieces"] + System.Environment.NewLine;
                    executionLogMessage += "Unit Count -" + dr["UnitCount"] + System.Environment.NewLine;
                    executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;

                    executionLogMessage += System.Environment.NewLine + "Please note, we have not processed this record.";

                    objCommon.WriteExecutionLog(executionLogMessage);
                    string fromMail = objCommon.GetConfigValue("FromMailID");
                    string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                    string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                    string toMail = objCommon.GetConfigValue("CDSSendBadDataEmailTo");
                    string subject = "DT-Cabinet Count and Unit Count found greater than 0 for this record, in file -" + fileName;
                    objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                    //dr.Delete();
                    itemsToDelete.Add(dr);
                    continue;
                }

                if (Convert.ToInt32(dr["Pieces"]) == 0 && Convert.ToInt32(dr["UnitCount"]) > 0)
                {
                    deliveryName = "Millwork";
                }

                deliveryName = deliveryName.Replace("'", "");

                int pieceunitcount = 0;
                if (Convert.ToInt32(dr["Pieces"]) > 0)
                {
                    pieceunitcount = Convert.ToInt32(dr["Pieces"]);
                }
                else if (Convert.ToInt32(dr["UnitCount"]) > 0)
                {
                    pieceunitcount = Convert.ToInt32(dr["UnitCount"]);
                }

                int miles = 0;
                if (!string.IsNullOrEmpty(Convert.ToString(dr["Miles"])))
                {
                    miles = Convert.ToInt32(dr["Miles"]);
                    if (miles < 0)
                    {
                        dr["Miles"] = 0;
                        miles = 0;
                    }
                }


                DataTable tblBillRatesFiltered = new DataTable();

                IEnumerable<DataRow> billratesfilteredRows = dtBillingRates.AsEnumerable()
                .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                         && row.Field<string>("DeliveryName") == deliveryName && row.Field<string>("IsActive") == "Y"
                         && (row.Field<decimal>("minimum_miles") <= miles)
                         && (miles <= row.Field<decimal>("maximum_miles")));

                if (billratesfilteredRows.Any())
                {
                    tblBillRatesFiltered = billratesfilteredRows.CopyToDataTable();
                }
                else
                {
                    billratesfilteredRows = dtBillingRates.AsEnumerable()
             .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
             && row.Field<string>("DeliveryName") is null && row.Field<string>("IsActive") == "Y"
               && (row.Field<decimal>("minimum_miles") <= miles)
                && (miles <= row.Field<decimal>("maximum_miles")));

                    if (billratesfilteredRows.Any())
                    {
                        tblBillRatesFiltered = billratesfilteredRows.CopyToDataTable();
                    }
                    else
                    {
                        //DataTable dtMissingConfData = new DataTable();
                        //dtMissingConfData = dtableOrderTemplateFinal.Select("[Customer Reference]= '" + dr["Customer Reference"] + "'").CopyToDataTable();

                        DataRow _newRow = dtMissingConfData.NewRow();
                        _newRow["Customer Reference"] = dr["Customer Reference"];
                        dtMissingConfData.Rows.Add(_newRow);
                        //dtMissingConfData.TableName = "MissingConf";
                        //string missingConfFilePath = objCommon.GetConfigValue("MissingConfFileFolder");
                        //objCommon.WriteDataToCsvFile(dtMissingConfData, missingConfFilePath, fileName, dateTime);

                        executionLogMessage = "DT-CDS Order Template File Generation Error" + System.Environment.NewLine;
                        executionLogMessage += "Missing Configuration for this record in billing rates" + System.Environment.NewLine;
                        executionLogMessage += "Customer Reference -" + dr["Customer Reference"] + System.Environment.NewLine;
                        executionLogMessage += "Cabinet Count -" + dr["Pieces"] + System.Environment.NewLine;
                        executionLogMessage += "Unit Count -" + dr["UnitCount"] + System.Environment.NewLine;
                        executionLogMessage += "Manufacturer -" + deliveryName + System.Environment.NewLine;
                        executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                        executionLogMessage += System.Environment.NewLine + "Please note, we have not processed this record.";

                        objCommon.WriteExecutionLog(executionLogMessage);
                        string fromMail = objCommon.GetConfigValue("FromMailID");
                        string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                        string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                        string toMail = objCommon.GetConfigValue("CDSSendMissingConfEmailTo");
                        string subject = "DT-Missing Configuration for this record in file - " + fileName;
                        objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                        itemsToDelete.Add(dr);
                        //  dr.Delete();
                        continue;
                    }
                }

                DataTable tblPayableRatesFiltered = new DataTable();
                IEnumerable<DataRow> payableratesfilteredRows;

                if (deliveryName != "Millwork")
                {
                    deliveryName = string.Empty;
                }

                if (deliveryType != null && deliveryType.ToUpper() != "WILL")
                {
                    if (deliveryName == "Millwork")
                    {
                        payableratesfilteredRows = dtPayableRates.AsEnumerable()
                       .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                       && (row.Field<string>("DeliveryName") == deliveryName)
                       && (row.Field<decimal>("minimumcount") <= pieceunitcount) && (pieceunitcount <= row.Field<decimal>("maximumcount")) && row.Field<string>("IsActive") == "Y"
                       && (row.Field<decimal>("minimum_miles") <= miles)
                       && (miles <= row.Field<decimal>("maximum_miles")));

                        if (payableratesfilteredRows.Any())
                        {
                            tblPayableRatesFiltered = payableratesfilteredRows.CopyToDataTable();
                        }
                        else
                        {
                            // No configuration found 

                            // need to copy the record into csv 

                            //DataTable dtMissingConfData = new DataTable();
                            //dtMissingConfData = dtableOrderTemplateFinal.Select("[Customer Reference]= '" + dr["Customer Reference"] + "'").CopyToDataTable();

                            DataRow _newRow = dtMissingConfData.NewRow();
                            _newRow["Customer Reference"] = dr["Customer Reference"];
                            dtMissingConfData.Rows.Add(_newRow);
                            //dtMissingConfData.TableName = "MissingConf";
                            //string missingConfFilePath = objCommon.GetConfigValue("MissingConfFileFolder");
                            //objCommon.WriteDataToCsvFile(dtMissingConfData, missingConfFilePath, fileName, dateTime);

                            executionLogMessage = "DT-CDS Order Template File Generation Error" + System.Environment.NewLine;
                            executionLogMessage += "Missing Configuration for this record in payable rates" + System.Environment.NewLine;
                            executionLogMessage += "Customer Reference -" + dr["Customer Reference"] + System.Environment.NewLine;
                            executionLogMessage += "Cabinet Count -" + dr["Pieces"] + System.Environment.NewLine;
                            executionLogMessage += "Unit Count -" + dr["UnitCount"] + System.Environment.NewLine;
                            executionLogMessage += "Manufacturer -" + deliveryName + System.Environment.NewLine;
                            executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                            executionLogMessage += System.Environment.NewLine + "Please note, we have not processed this record.";

                            objCommon.WriteExecutionLog(executionLogMessage);
                            string fromMail = objCommon.GetConfigValue("FromMailID");
                            string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                            string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                            string toMail = objCommon.GetConfigValue("CDSSendMissingConfEmailTo");
                            string subject = "DT-Missing Configuration for this record in file - " + fileName;
                            objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                            itemsToDelete.Add(dr);
                            //  dr.Delete();
                            continue;

                        }
                    }
                    else
                    {
                        //  DataTable tblPayableRatesFiltered = new DataTable();
                        //  IEnumerable<DataRow> payableratesfilteredRows = dtPayableRates.AsEnumerable()

                        payableratesfilteredRows = dtPayableRates.AsEnumerable()
                        .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                        && row.Field<string>("DeliveryName") == deliveryName && row.Field<string>("IsActive") == "Y"
                        && (row.Field<decimal>("minimum_miles") <= miles)
                        && (miles <= row.Field<decimal>("maximum_miles")));

                        if (payableratesfilteredRows.Any())
                        {
                            tblPayableRatesFiltered = payableratesfilteredRows.CopyToDataTable();
                        }
                        else
                        {
                            payableratesfilteredRows = dtPayableRates.AsEnumerable()
                        .Where(row => (row.Field<DateTime>("EffectiveStartDate") <= dtdeliveryDate) && (dtdeliveryDate <= row.Field<DateTime>("EffectiveEndDate"))
                        && row.Field<string>("DeliveryName") is null && row.Field<string>("IsActive") == "Y"
                        && (row.Field<decimal>("minimum_miles") <= miles)
                        && (miles <= row.Field<decimal>("maximum_miles")));

                            if (payableratesfilteredRows.Any())
                            {
                                tblPayableRatesFiltered = payableratesfilteredRows.CopyToDataTable();
                            }
                            else
                            {
                                // No configuration found 
                                // need to copy the record into csv 
                                //DataTable dtMissingConfData = new DataTable();
                                //dtMissingConfData = dtableOrderTemplateFinal.Select("[Customer Reference]= '" + dr["Customer Reference"] + "'").CopyToDataTable();

                                DataRow _newRow = dtMissingConfData.NewRow();
                                _newRow["Customer Reference"] = dr["Customer Reference"];
                                dtMissingConfData.Rows.Add(_newRow);
                                //dtMissingConfData.TableName = "MissingConf";
                                //string missingConfFilePath = objCommon.GetConfigValue("MissingConfFileFolder");
                                //objCommon.WriteDataToCsvFile(dtMissingConfData, missingConfFilePath, fileName, dateTime);

                                executionLogMessage = "DT-CDS Order Template File Generation Error" + System.Environment.NewLine;
                                executionLogMessage += "Missing Configuration for this record in payable rates" + System.Environment.NewLine;
                                executionLogMessage += "Customer Reference -" + dr["Customer Reference"] + System.Environment.NewLine;
                                executionLogMessage += "Cabinet Count -" + dr["Pieces"] + System.Environment.NewLine;
                                executionLogMessage += "Unit Count -" + dr["UnitCount"] + System.Environment.NewLine;
                                executionLogMessage += "Manufacturer -" + deliveryName + System.Environment.NewLine;
                                executionLogMessage += "FileName -" + fileName + System.Environment.NewLine;
                                executionLogMessage += System.Environment.NewLine + "Please note, we have not processed this record.";

                                objCommon.WriteExecutionLog(executionLogMessage);
                                string fromMail = objCommon.GetConfigValue("FromMailID");
                                string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                string toMail = objCommon.GetConfigValue("CDSSendMissingConfEmailTo");
                                string subject = "DT-Missing Configuration for this record in file - " + fileName;
                                objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                                itemsToDelete.Add(dr);
                                //  dr.Delete();
                                continue;

                            }
                        }
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
                        executionLogMessage = "DT-Diesel price is missing for date  " + dtdeliveryDate.ToShortDateString() + System.Environment.NewLine;
                        executionLogMessage += "So not able to process this file, please update the fsc sheet with appropriate values." + System.Environment.NewLine;
                        executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                        objCommon.WriteExecutionLog(executionLogMessage);
                        string fromMail = objCommon.GetConfigValue("FromMailID");
                        string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                        string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                        string toMail = objCommon.GetConfigValue("CDSSendMissingDieselPriceEmailTo");
                        string subject = "DT-Diesel price is missing for date  " + dtdeliveryDate.ToShortDateString();
                        objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                        throw new NullReferenceException("Diesel price is missing for date  " + dtdeliveryDate.ToShortDateString());
                    }

                    if (dtFSCRatesfromDB.Rows.Count > 0)
                    {
                        IEnumerable<DataRow> fscratesfilteredRows = dtFSCRatesfromDB.AsEnumerable()
                        .Where(row => (row.Field<decimal>("Start") <= fuelcharge) && (fuelcharge <= row.Field<decimal>("Stop"))
                         && row.Field<string>("DeliveryName") == deliveryName && row.Field<string>("IsActive") == "Y");

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
                            && row.Field<string>("DeliveryName") is null && row.Field<string>("IsActive") == "Y");
                            if (fscratesfilteredRows.Any())
                            {
                                tblFSCRatesFiltered = fscratesfilteredRows.CopyToDataTable();
                                fscratePercentage = Convert.ToDouble(tblFSCRatesFiltered.Rows[0]["Percent_FSCAMount"]);
                                fscratetype = Convert.ToString(tblFSCRatesFiltered.Rows[0]["Type"]);
                            }
                            else
                            {
                                executionLogMessage = "DT-FSC Billing Rates missing for this fuel charge   " + fuelcharge + System.Environment.NewLine;
                                executionLogMessage += "So not able to process this file, please update the billable rates mapping table with appropriate values";
                                executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                objCommon.WriteExecutionLog(executionLogMessage);
                                string fromMail = objCommon.GetConfigValue("FromMailID");
                                string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                string toMail = objCommon.GetConfigValue("CDSSendMissingDieselPriceEmailTo");
                                string subject = "DT-FSC Billing Rates missing for this fuel charge   " + fuelcharge;
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
                        && row.Field<string>("DeliveryName") == deliveryName && row.Field<string>("IsActive") == "Y");

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
                        && row.Field<string>("DeliveryName") is null && row.Field<string>("IsActive") == "Y");
                            if (CarrierfscratesfilteredRows.Any())
                            {
                                tblCarrierFSCRatesFiltered = CarrierfscratesfilteredRows.CopyToDataTable();
                                carrierfscratePercentage = Convert.ToDouble(tblCarrierFSCRatesFiltered.Rows[0]["Percent_FSCAMount"]);
                                carrierfscratetype = Convert.ToString(tblCarrierFSCRatesFiltered.Rows[0]["Type"]);
                            }
                            else
                            {
                                executionLogMessage = "DT-FSC Payable Rates missing for this fuel charge   " + fuelcharge + System.Environment.NewLine;
                                executionLogMessage += "So not able to process this file, please update the payable rates mapping table with appropriate values";
                                executionLogMessage += "Found exception while processing the file, filename  -" + fileName + System.Environment.NewLine;
                                objCommon.WriteExecutionLog(executionLogMessage);
                                string fromMail = objCommon.GetConfigValue("FromMailID");
                                string fromPassword = objCommon.GetConfigValue("FromMailPasssword");
                                string disclaimer = objCommon.GetConfigValue("MailDisclaimer");
                                string toMail = objCommon.GetConfigValue("CDSSendMissingDieselPriceEmailTo");
                                string subject = "DT-FSC Billing Rates missing for this fuel charge   " + fuelcharge;
                                objCommon.SendMail(fromMail, fromPassword, disclaimer, toMail, "", subject, executionLogMessage, "");
                                throw new NullReferenceException("Diesel prices missing for date  " + dtdeliveryDate);
                            }
                        }
                    }
                }

                const double daysToMonths = 30.4368499;

                if (tblBillRatesFiltered.Rows.Count > 0)
                {
                    double billRate = 0;
                    double minimumRate = 0;
                    double parts_only_charge = 0;
                    string storage_charge_type = string.Empty;
                    string storage_charge_unit = string.Empty;
                    int storage_charge_minimum_days = 0;
                    string stair_charge_type = string.Empty;
                    int stair_number_of_minimum_flights = 0;
                    int numberofStairs = 0;

                    int numberofdiffDays = 0;
                    double totalStorageCharge = 0;
                    double totalStaireCharge = 0;
                    int numberofflight = 0;
                    double maximum_miles = 0;
                    double mileage_charge_rate = 0;
                    double maximum_men = 0;
                    double extra_man_fees = 0;
                    string will_call_type = string.Empty;
                    double will_call_charge = 0;
                    // int miles = 0;
                    int men = 0;
                    double totalMilesCharge = 0;
                    double totalExtraManFee = 0;


                    double add_charge_amt1 = 0;
                    double add_charge_amt2 = 0;

                    double minimumCount = 0;

                    string return_rate_type = string.Empty;
                    double return_rate = 0;


                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["minimum_rate"])))
                    {
                        minimumRate = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["minimum_rate"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["minimumcount"])))
                    {
                        minimumCount = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["minimumcount"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["parts_only_charge"])))
                    {
                        parts_only_charge = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["parts_only_charge"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["storage_charge_type"])))
                    {
                        storage_charge_type = Convert.ToString(tblBillRatesFiltered.Rows[0]["storage_charge_type"]); ;
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["storage_charge_unit"])))
                    {
                        storage_charge_unit = Convert.ToString(tblBillRatesFiltered.Rows[0]["storage_charge_unit"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["storage_charge_minimum_days"])))
                    {
                        storage_charge_minimum_days = Convert.ToInt32(tblBillRatesFiltered.Rows[0]["storage_charge_minimum_days"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["stair_charge_type"])))
                    {
                        stair_charge_type = Convert.ToString(tblBillRatesFiltered.Rows[0]["stair_charge_type"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["stair_number_of_minimum_flights"])))
                    {
                        stair_number_of_minimum_flights = Convert.ToInt32(tblBillRatesFiltered.Rows[0]["stair_number_of_minimum_flights"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["add_charge_amt1"])))
                    {
                        add_charge_amt1 = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["add_charge_amt1"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["add_charge_amt2"])))
                    {
                        add_charge_amt2 = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["add_charge_amt2"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Stairs"])))
                    {
                        numberofStairs = Convert.ToInt32(dr["Stairs"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["maximum_miles"])))
                    {
                        maximum_miles = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["maximum_miles"]);
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["mileage_charge_rate"])))
                    {
                        mileage_charge_rate = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["mileage_charge_rate"]);
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["maximum_men"])))
                    {
                        maximum_men = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["maximum_men"]);
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["extra_man_fees"])))
                    {
                        extra_man_fees = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["extra_man_fees"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["will_call_type"])))
                    {
                        will_call_type = Convert.ToString(tblBillRatesFiltered.Rows[0]["will_call_type"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["will_call_charge"])))
                    {
                        will_call_charge = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["will_call_charge"]);
                    }


                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["return_rate_type"])))
                    {
                        return_rate_type = Convert.ToString(tblBillRatesFiltered.Rows[0]["return_rate_type"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["return_rate"])))
                    {
                        return_rate = Convert.ToDouble(tblBillRatesFiltered.Rows[0]["return_rate"]);
                    }


                    if (diffDays > storage_charge_minimum_days)
                    {
                        numberofdiffDays = diffDays - storage_charge_minimum_days;
                    }

                    if (numberofStairs > stair_number_of_minimum_flights)
                    {
                        numberofflight = numberofStairs - stair_number_of_minimum_flights;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Miles"])))
                    {
                        miles = Convert.ToInt32(dr["Miles"]);
                        if (miles < 0)
                        {
                            dr["Miles"] = 0;
                            miles = 0;
                        }
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Men"])))
                    {
                        men = Convert.ToInt32(dr["Men"]);
                    }

                    if (miles > maximum_miles)
                    {
                        miles = miles - Convert.ToInt32(maximum_miles);

                        totalMilesCharge = Math.Round(Convert.ToDouble(miles * mileage_charge_rate), 2);

                    }

                    if (men > maximum_men)
                    {
                        men = men - Convert.ToInt32(maximum_men);

                        totalExtraManFee = Math.Round(Convert.ToDouble(men * extra_man_fees), 2);
                    }


                    if (Convert.ToInt32(dr["Pieces"]) == 0 && Convert.ToInt32(dr["UnitCount"]) == 0)
                    {
                        billRate = parts_only_charge;
                        dr["Bill Rate"] = Math.Round(Convert.ToDouble(billRate), 2);
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"])))
                        {
                            if (pieceunitcount < minimumCount)
                            {
                                billRate = minimumRate;
                            }
                            else
                            {
                                billRate = minimumRate + (Convert.ToDouble(pieceunitcount - minimumCount) * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]));
                            }
                            //billRate = Convert.ToDouble(pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt1"]));
                            //if (billRate < minimumRate)
                            //{
                            //    billRate = minimumRate;
                            //}
                            dr["Bill Rate"] = Math.Round(Convert.ToDouble(billRate), 2);

                            // need to implement a minimum count column logic
                        }
                        // Calcuation of Total Storage Amount
                        if (storage_charge_type.ToUpper() == "CABINET")
                        {
                            if (storage_charge_unit.ToUpper() == "DAY")
                            {
                                totalStorageCharge = numberofdiffDays * add_charge_amt2 * Convert.ToDouble(pieceunitcount);
                            }
                            else if (storage_charge_unit.ToUpper() == "MONTH")
                            {
                                // var totalMonths = Math.Round((numberofdiffDays % 365) / 30);
                                var totalMonths = Convert.ToDouble(Math.Round(Convert.ToDouble(numberofdiffDays / daysToMonths)));
                                // var totalMonths = Convert.ToDouble(Math.Round(Convert.ToDouble((numberofdiffDays % 365) / 30)));
                                totalStorageCharge = totalMonths * add_charge_amt2 * Convert.ToDouble(pieceunitcount);
                            }
                        }
                        else if (storage_charge_type.ToUpper() == "ORDER")
                        {
                            // totalStorageCharge = numberofdiffDays * rate_buck_amt2 * 1;
                            if (storage_charge_unit.ToUpper() == "DAY")
                            {
                                totalStorageCharge = numberofdiffDays * add_charge_amt2 * 1;
                            }
                            else if (storage_charge_unit.ToUpper() == "MONTH")
                            {
                                var totalMonths = Convert.ToDouble(Math.Round(Convert.ToDouble(numberofdiffDays / daysToMonths)));
                                totalStorageCharge = totalMonths * add_charge_amt2 * 1;
                            }
                        }

                        // Calcuation of Total Stair Charge
                        if (stair_charge_type.ToUpper() == "CABINET")
                        {
                            totalStaireCharge = numberofflight * add_charge_amt1 * Convert.ToDouble(pieceunitcount);
                        }
                        else if (stair_charge_type.ToUpper() == "ORDER")
                        {
                            totalStaireCharge = numberofflight * add_charge_amt1 * 1;
                        }

                    }

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

                    if (deliveryType != null)
                    {
                        if (deliveryType.ToUpper() == "WILL")
                        {
                            dr["Correct Driver Number"] = objCommon.GetConfigValue("CDSCorrectDriverNumberforWill");

                            if (!string.IsNullOrEmpty(Convert.ToString(will_call_charge)))
                            {
                                if (will_call_type.ToString().ToUpper() == "F")
                                {
                                    billRate = will_call_charge;
                                }
                                else
                                {
                                    billRate = Math.Round(Convert.ToDouble(billRate * will_call_charge) / 100, 2);
                                }
                                dr["Bill Rate"] = Math.Round(Convert.ToDouble(billRate), 2);
                                //  totalStorageCharge = 0;
                                totalStaireCharge = 0;
                                totalMilesCharge = 0;
                                totalExtraManFee = 0;
                                dr["FSC"] = 0;
                            }
                        }
                        else if (deliveryType.ToUpper() == "RETURN")
                        {
                            if (!string.IsNullOrEmpty(Convert.ToString(return_rate)))
                            {
                                if (return_rate_type.ToString().ToUpper() == "F")
                                {
                                    billRate = return_rate;
                                }
                                else
                                {
                                    billRate = Math.Round(Convert.ToDouble(billRate * return_rate) / 100, 2);
                                }
                                dr["Bill Rate"] = Math.Round(Convert.ToDouble(billRate), 2);
                                totalStorageCharge = 0;
                                //totalStaireCharge = 0;
                                //totalMilesCharge = 0;
                                //totalExtraManFee = 0;
                                //dr["FSC"] = 0;
                            }
                        }
                    }



                    //dr["rate_buck_amt2"] = totalStorageCharge;
                    //dr["Pieces ACC"] = totalStaireCharge;
                    //dr["Miles"] = Convert.ToInt32(totalMilesCharge);
                    //dr["rate_buck_amt4"] = (totalExtraManFee);

                    dr["add_charge_amt1"] = totalStaireCharge;
                    dr["add_charge_amt2"] = totalStorageCharge;
                    dr["add_charge_amt3"] = Convert.ToInt32(totalMilesCharge);
                    dr["add_charge_amt4"] = totalExtraManFee;

                    if (pieceunitcount > 0)
                    {

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"])))
                            dr["rate_buck_amt2"] = pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"]);

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"])))
                            dr["rate_buck_amt4"] = pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt4"]);

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"])))
                            dr["rate_buck_amt5"] = pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt5"]);

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"])))
                            dr["rate_buck_amt6"] = pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt6"]);

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"])))
                            dr["rate_buck_amt7"] = pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt7"]);

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"])))
                            dr["rate_buck_amt8"] = pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt8"]);

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"])))
                            dr["rate_buck_amt9"] = pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt9"]);

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"])))
                            dr["rate_buck_amt11"] = pieceunitcount * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt11"]);

                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"])))
                            dr["rate_buck_amt2"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt2"]);

                        if (!string.IsNullOrEmpty(Convert.ToString(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"])))
                            dr["Pieces ACC"] = 1 * Convert.ToDouble(tblBillRatesFiltered.Rows[0]["rate_buck_amt3"]);

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

                }

                if (tblPayableRatesFiltered.Rows.Count > 0)
                {

                    double carrierBasePay = 0;
                    double minimumRate = 0;
                    double parts_only_charge = 0;
                    //string storage_charge_type = string.Empty;
                    //string storage_charge_unit = string.Empty;
                    //int storage_charge_minimum_days = 0;
                    string stair_charge_type = string.Empty;
                    int stair_number_of_minimum_flights = 0;
                    int numberofStairs = 0;
                    //double charge1 = 0;
                    //double charge2 = 0;
                    //double charge3 = 0;

                    // int numberofdiffDays = 0;
                    // double totalStorageCharge = 0;
                    double totalStaireCharge = 0;
                    int numberofflight = 0;
                    double maximum_miles = 0;
                    double mileage_charge_rate = 0;
                    double maximum_men = 0;
                    double extra_man_fees = 0;
                    string will_call_type = string.Empty;
                    double will_call_charge = 0;
                    //  int miles = 0;
                    int men = 0;
                    double totalMilesCharge = 0;
                    double totalExtraManFee = 0;
                    //  string deliveryType = string.Empty;

                    double add_charge_amt1 = 0;
                    double add_charge_amt2 = 0;

                    string return_rate_type = string.Empty;
                    double return_rate = 0;

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["minimum_rate"])))
                    {
                        minimumRate = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["minimum_rate"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["parts_only_charge"])))
                    {
                        parts_only_charge = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["parts_only_charge"]); ;
                    }


                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["stair_charge_type"])))
                    {
                        stair_charge_type = Convert.ToString(tblPayableRatesFiltered.Rows[0]["stair_charge_type"]); ;
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["stair_number_of_minimum_flights"])))
                    {
                        stair_number_of_minimum_flights = Convert.ToInt32(tblPayableRatesFiltered.Rows[0]["stair_number_of_minimum_flights"]); ;
                    }


                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["add_charge_amt1"])))
                    {
                        add_charge_amt1 = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["add_charge_amt1"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["add_charge_amt2"])))
                    {
                        add_charge_amt2 = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["add_charge_amt2"]);
                    }


                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Stairs"])))
                    {
                        numberofStairs = Convert.ToInt32(dr["Stairs"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["maximum_miles"])))
                    {
                        maximum_miles = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["maximum_miles"]);
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["mileage_charge_rate"])))
                    {
                        mileage_charge_rate = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["mileage_charge_rate"]);
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["maximum_men"])))
                    {
                        maximum_men = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["maximum_men"]);
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["extra_man_fees"])))
                    {
                        extra_man_fees = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["extra_man_fees"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["will_call_type"])))
                    {
                        will_call_type = Convert.ToString(tblPayableRatesFiltered.Rows[0]["will_call_type"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["will_call_charge"])))
                    {
                        will_call_charge = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["will_call_charge"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["return_rate_type"])))
                    {
                        return_rate_type = Convert.ToString(tblPayableRatesFiltered.Rows[0]["return_rate_type"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["return_rate"])))
                    {
                        return_rate = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["return_rate"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Delivery Type"])))
                    {
                        deliveryType = Convert.ToString(dr["Delivery Type"]);
                    }


                    //if (diffDays > storage_charge_minimum_days)
                    //{
                    //    numberofdiffDays = diffDays - storage_charge_minimum_days;
                    //}

                    if (numberofStairs > stair_number_of_minimum_flights)
                    {
                        numberofflight = numberofStairs - stair_number_of_minimum_flights;
                    }


                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Miles"])))
                    {
                        miles = Convert.ToInt32(dr["Miles"]);
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Men"])))
                    {
                        men = Convert.ToInt32(dr["Men"]);
                    }

                    if (miles > maximum_miles)
                    {
                        miles = miles - Convert.ToInt32(maximum_miles);

                        totalMilesCharge = Math.Round(Convert.ToDouble(miles * mileage_charge_rate), 2);

                    }

                    if (men > maximum_men)
                    {
                        men = men - Convert.ToInt32(maximum_men);

                        totalExtraManFee = Math.Round(Convert.ToDouble(men * extra_man_fees), 2);
                    }

                    if (Convert.ToInt32(dr["Pieces"]) == 0 && Convert.ToInt32(dr["UnitCount"]) == 0)
                    {
                        carrierBasePay = parts_only_charge;
                        dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);
                    }
                    else
                    {

                        if (deliveryName == "Millwork")
                        {
                            //carrierBasePay = Convert.ToDouble(Convert.ToDouble(pieceunitcount) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]));
                            carrierBasePay = Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);
                            if (carrierBasePay < minimumRate)
                            {
                                carrierBasePay = minimumRate;
                            }
                            dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);

                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                            {
                                carrierBasePay = Convert.ToDouble(Convert.ToDouble(pieceunitcount) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]));
                                if (carrierBasePay < minimumRate)
                                {
                                    carrierBasePay = minimumRate;
                                }
                                dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);
                            }

                        }



                        // Calcuation of Total Stair Charge
                        if (stair_charge_type.ToUpper() == "CABINET")
                        {
                            totalStaireCharge = numberofflight * add_charge_amt1 * Convert.ToDouble(pieceunitcount);
                        }
                        else if (stair_charge_type.ToUpper() == "ORDER")
                        {
                            totalStaireCharge = numberofflight * add_charge_amt1 * 1;
                        }
                    }
                    if (!string.IsNullOrEmpty(Convert.ToString(carrierfscratePercentage)))
                    {
                        if (fscratetype.ToString().ToUpper() == "F")
                        {
                            dr["Carrier FSC"] = Math.Round(Convert.ToDouble(carrierfscratePercentage), 2);
                        }
                        else
                        {
                            dr["Carrier FSC"] = Math.Round(Convert.ToDouble(carrierBasePay * carrierfscratePercentage) / 100, 2);
                        }
                    }

                    if (deliveryType != null)
                    {
                        if (deliveryType.ToUpper() == "WILL")
                        {
                            if (!string.IsNullOrEmpty(Convert.ToString(will_call_charge)))
                            {
                                if (will_call_type.ToString().ToUpper() == "F")
                                {
                                    carrierBasePay = will_call_charge;
                                }
                                else
                                {
                                    carrierBasePay = Math.Round(Convert.ToDouble(carrierBasePay * will_call_charge) / 100, 2);
                                }
                                dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);
                                // totalStorageCharge = 0;
                                totalStaireCharge = 0;
                                totalMilesCharge = 0;
                                totalExtraManFee = 0;
                                dr["Carrier FSC"] = 0;
                            }
                        }
                        else if (deliveryType.ToUpper() == "RETURN")
                        {
                            if (!string.IsNullOrEmpty(Convert.ToString(return_rate)))
                            {
                                if (return_rate_type.ToString().ToUpper() == "F")
                                {
                                    carrierBasePay = return_rate;
                                }
                                else
                                {
                                    carrierBasePay = Math.Round(Convert.ToDouble(carrierBasePay * return_rate) / 100, 2);
                                }
                                dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);
                                //totalStorageCharge = 0;
                                //totalStaireCharge = 0;
                                //totalMilesCharge = 0;
                                //totalExtraManFee = 0;
                                //dr["FSC"] = 0;
                            }
                        }
                    }



                    //  dr["charge2"] = totalStorageCharge;
                    // dr["charge3"] = totalStaireCharge;
                    // dr["Miles"] = Convert.ToInt32(totalMilesCharge);
                    // dr["charge4"] = (totalExtraManFee);

                    dr["charge2"] = totalStaireCharge;
                    dr["charge3"] = totalMilesCharge;
                    dr["charge4"] = totalExtraManFee;

                    //if (!string.IsNullOrEmpty(Convert.ToString(carrierfscratePercentage)))
                    //{
                    //    if (carrierfscratetype.ToString().ToUpper() == "F")
                    //    {
                    //        dr["Carrier FSC"] = Math.Round(Convert.ToDouble(carrierfscratePercentage), 2);
                    //    }
                    //    else
                    //    {
                    //        dr["Carrier FSC"] = Math.Round(Convert.ToDouble(carrierBasePay * carrierfscratePercentage) / 100, 2);
                    //    }
                    //}

                    if (pieceunitcount > 0)
                    {
                        //if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                        //{
                        //    carrierBasePay = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);
                        //    if (carrierBasePay < minimumRate)
                        //    {
                        //        carrierBasePay = minimumRate;
                        //    }
                        //    dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);
                        //}

                        if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge5"])))
                            dr["Carrier ACC"] = pieceunitcount * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge5"]);

                    }
                    else
                    {
                        //if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge1"])))
                        //{
                        //    dr["Carrier Base Pay"] = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);
                        //    carrierBasePay = Convert.ToDouble(dr["Pieces"]) * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge1"]);
                        //    if (carrierBasePay < minimumRate)
                        //    {
                        //        carrierBasePay = minimumRate;
                        //    }
                        //    dr["Carrier Base Pay"] = Math.Round(Convert.ToDouble(carrierBasePay), 2);
                        //}
                        if (!string.IsNullOrEmpty(Convert.ToString(tblPayableRatesFiltered.Rows[0]["charge5"])))
                            dr["Carrier ACC"] = 1 * Convert.ToDouble(tblPayableRatesFiltered.Rows[0]["charge5"]);


                    }
                }



            }

            if (itemsToDelete.Count > 0)
            {
                if (dtBadData.Rows.Count > 0)
                {
                    var baditemsToDelete = new List<DataRow>();

                    foreach (DataRow dr in dtCDSBadData.Rows)
                    {
                        // DataTable dtBad = dtBadData.Select("[Customer Reference]= '" + dr["CDS#"] + "'").CopyToDataTable();
                        DataRow[] drBadresult = dtBadData.Select("[Customer Reference]= '" + dr[customerrefmappingcolumnname] + "'");

                        if (!drBadresult.Any())
                        {
                            baditemsToDelete.Add(dr);
                        }
                    }
                    foreach (var item in baditemsToDelete)
                    {
                        dtCDSBadData.Rows.Remove(item);
                    }

                    dtCDSBadData.AcceptChanges();

                    dtCDSBadData.TableName = "BadData";
                    string badDataFilePath = objCommon.GetConfigValue("BadDataFileFolder");
                    objCommon.WriteDataToCsvFile(dtCDSBadData, badDataFilePath, fileName, dateTime);

                }

                if (dtMissingConfData.Rows.Count > 0)
                {
                    var missingconfigitemsToDelete = new List<DataRow>();

                    foreach (DataRow dr in dtCDSMissingConfigData.Rows)
                    {
                        //DataTable dtmissingdt = dtMissingConfData.Select("[Customer Reference]= '" + dr["CDS#"] + "'").CopyToDataTable();
                        DataRow[] drmissingresult = dtMissingConfData.Select("[Customer Reference]= '" + dr[customerrefmappingcolumnname] + "'");
                        if (!drmissingresult.Any())
                        {
                            missingconfigitemsToDelete.Add(dr);
                        }
                    }

                    foreach (var item in missingconfigitemsToDelete)
                    {
                        dtCDSMissingConfigData.Rows.Remove(item);
                    }

                    dtCDSMissingConfigData.AcceptChanges();

                    dtCDSMissingConfigData.TableName = "MissingConf";
                    string missingConfFilePath = objCommon.GetConfigValue("MissingConfFileFolder");
                    objCommon.WriteDataToCsvFile(dtCDSMissingConfigData, missingConfFilePath, fileName, dateTime);

                }

                // to delete the row from the dtableOrderTemplateFinal if found bad data or missing config data.
                foreach (var item in itemsToDelete)
                {
                    dtableOrderTemplateFinal.Rows.Remove(item);
                }
            }
            dtableOrderTemplateFinal.AcceptChanges();
            return returnResponse;
        }

    }
}
