using Microsoft.ApplicationBlocks.Data;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace JWLRetriveEmail
{
    class clsEnterpriseDB : clsCommon
    {
        public DSResponse GetReadEmailMappingDetails_Enterprise(string CustomerName, string LocationCode, string ProductCode)
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

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection_Enterprise"), CommandType.StoredProcedure, "USP_S_READEMAIL_CUSTOMERMAPPING",
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
                WriteErrorLog(ex, "GetReadEmailMappingDetails_Enterprise");
            }
            return objResponse;
        }

        public DSResponse GetOrderPostTemplateDetails_Enterprise(string CustomerName, string LocationCode, string ProductCode, string ProductSubCode)
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

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection_Enterprise"), CommandType.StoredProcedure, "USP_S_OrderPostTemplate_CustomerMappingColumns",
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
                WriteErrorLog(ex, "GetOrderPostTemplateDetails_Enterprise");
            }
            return objResponse;
        }

        public DSResponse GetBillingRatesAndPayableRates_CustomerMappingDetails_Enterprise(string company, string customerNumber)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramCompany = new SqlParameter("@Company", SqlDbType.Int);
                paramCompany.Value = company;

                SqlParameter paramCustomerNumber = new SqlParameter("@CustomerNumber", SqlDbType.Int);
                paramCustomerNumber.Value = customerNumber;

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection_Enterprise"), CommandType.StoredProcedure, "USP_S_BillingRates_PayableRates_CustomerMapping",
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
                WriteErrorLog(ex, "GetBillingRatesAndPayableRates_CustomerMappingDetails_Enterprise");
            }
            return objResponse;
        }
       
        public DSResponse GetFSCRates_MappingDetails_Enterprise(int company, int customerNumber)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramCompany = new SqlParameter("@Company", SqlDbType.Int);
                paramCompany.Value = company;

                SqlParameter paramCustomerNumber = new SqlParameter("@CustomerNumber", SqlDbType.Int);
                paramCustomerNumber.Value = customerNumber;

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection_Enterprise"), CommandType.StoredProcedure, "USP_S_FSCRate_Mapping",
                    paramCompany, paramCustomerNumber);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "FSC Rates customer mapping details not found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetFSCRates_MappingDetails_Enterprise");
            }
            return objResponse;
        }

        public DSResponse GetStoreBand_DeficitWeightRatingDetails_Enterprise(int company, int customerNumber)
        {
            DSResponse objResponse = new DSResponse();
            try
            {
                DataSet dsDtls = new DataSet();

                SqlParameter paramCompany = new SqlParameter("@Company", SqlDbType.Int);
                paramCompany.Value = company;

                SqlParameter paramCustomerNumber = new SqlParameter("@CustomerNumber", SqlDbType.Int);
                paramCustomerNumber.Value = customerNumber;

                dsDtls = SqlHelper.ExecuteDataset(GetConfigValue("DBConnection_Enterprise"), CommandType.StoredProcedure, "USP_S_StoreBand_DeficitWeightRating_Mapping",
                    paramCompany, paramCustomerNumber);
                if (dsDtls.Tables[0].Rows.Count > 0)
                {
                    objResponse.DS = dsDtls;
                    objResponse.dsResp.ResponseVal = true;
                }
                else
                {
                    objResponse.dsResp.ResponseVal = false;
                    objResponse.dsResp.Reason = "Deficit Weight Rating details not found";
                }
            }
            catch (Exception ex)
            {
                objResponse.dsResp.ResponseVal = false;
                WriteErrorLog(ex, "GetStoreBand_DeficitWeightRatingDetails_Enterprise");
            }
            return objResponse;
        }
    }
}
