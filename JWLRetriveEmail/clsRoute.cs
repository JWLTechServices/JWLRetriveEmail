using Microsoft.ApplicationBlocks.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace JWLRetriveEmail
{
    class clsRoute : clsCommon
    {
        public ReturnResponse DataTracRouteStopGetAPI(string UniqueId)
        {
            ReturnResponse objresponse = new ReturnResponse();

            string json = string.Empty;
            clsCommon objCommon = new clsCommon();
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(5);
                    // string url = objCommon.GetConfigValue("DatatracURL") + "/route_stop";
                    string url = objCommon.GetConfigValue("DatatracURL") + "/route_stop/" + UniqueId;
                    client.DefaultRequestHeaders
                      .Accept
                      .Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var Username = objCommon.GetConfigValue("DatatracUserName");
                    var Password = objCommon.GetConfigValue("DatatracPassword");

                    UTF8Encoding utf8 = new UTF8Encoding();

                    byte[] encodedBytes = utf8.GetBytes(Username + ":" + Password);
                    string userCredentialsEncoding = Convert.ToBase64String(encodedBytes);
                    client.DefaultRequestHeaders.Add("Authorization", "Basic " + userCredentialsEncoding);
                    //  JObject jsonobj = JObject.Parse(jsonreq);
                    //  string payload = jsonobj.ToString();
                    //string payload = JsonConvert.SerializeObject(objheaderdetails);
                    var response = client.GetAsync(url).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        objresponse.Reason = response.Content.ReadAsStringAsync().Result;
                        objresponse.ResponseVal = true;
                    }
                    else
                    {
                        objresponse.ResponseVal = false;
                        objresponse.Reason = response.Content.ReadAsStringAsync().Result;
                        string strExecutionLogMessage = "Datatrac response failed for the customer reference number ";// + objorderdetails.order.reference + System.Environment.NewLine;
                        strExecutionLogMessage += "UniqueId:" + UniqueId + System.Environment.NewLine;
                        strExecutionLogMessage += "Response:" + objresponse.Reason;
                        objCommon.WriteExecutionLog(strExecutionLogMessage);

                    }

                }
            }
            catch (Exception ex)
            {
                string strExecutionLogMessage = "exception in CallDataTracRouteStopPutAPI " + ex;
                objresponse.Reason = strExecutionLogMessage;
                objresponse.ResponseVal = false;
                objCommon.WriteExecutionLog(strExecutionLogMessage);
                objCommon.WriteErrorLog(ex, strExecutionLogMessage);
            }
            return objresponse;
        }
        public string GenerateUniqueNumber(int company_number, int control_number, string orderSettlement_UniqueIdSuffix = "")
        {
            string result;
            // result = company_number.ToString().PadLeft(3, '0') + "0" + control_number + "0";
            result = company_number.ToString().PadLeft(3, '0') + control_number.ToString().PadLeft(8, '0') + "0";
            if (orderSettlement_UniqueIdSuffix != "")
            {
                result = result + orderSettlement_UniqueIdSuffix; // "D1";   //get it from excel future mapping
            }
            return result;
        }


    }
    public class ErrorResponse
    {
        public string error { get; set; }
        public string status { get; set; }
        public string code { get; set; }
        public string reference { get; set; }

    }




    public class items
    {
        public int company_number { get; set; }
        public int unique_id { get; set; }
        public string actual_cod_type { get; set; }
        public string barcodes_unique { get; set; }
        public string cod_type { get; set; }
        public string photos_exist { get; set; }
        public string @return { get; set; }
        public int expected_pieces { get; set; }
        public int expected_weight { get; set; }
        public string item_description { get; set; }
        public string item_number { get; set; }
        public string container_id { get; set; }
        public string reference { get; set; }

    }
    public class ResponseRouteStop
    {
        public object room { get; set; }
        public object unique_id { get; set; }
        public object c2_paperwork { get; set; }
        public object company_number_text { get; set; }
        public object company_number { get; set; }
        public object addl_charge_code11 { get; set; }
        public object billing_override_amt { get; set; }
        public object addl_charge_occur1 { get; set; }
        public object updated_time { get; set; }
        public object stop_sequence { get; set; }
        public object phone { get; set; }
        public object city { get; set; }
        public object created_by { get; set; }
        public List<object> signature_images { get; set; }
        public object pricing_zone { get; set; }
        public object signature_filename { get; set; }
        public object addl_charge_code10 { get; set; }
        public object cod_check_no { get; set; }
        public object length { get; set; }
        public object expected_weight { get; set; }
        public object actual_settlement_amt { get; set; }
        public object actual_pieces { get; set; }
        public object updated_date { get; set; }
        public object schedule_stop_id { get; set; }
        public object photos_exist { get; set; }
        public object stop_type_text { get; set; }
        public object stop_type { get; set; }
        public object @return { get; set; }
        public object addl_charge_code6 { get; set; }
        public object dispatch_zone { get; set; }
        public object upload_time { get; set; }
        public object actual_cod_amt { get; set; }
        public object location_accuracy { get; set; }
        public List<Progress> progress { get; set; }
        public object received_route { get; set; }
        public object override_settle_percent { get; set; }
        public object cod_amount { get; set; }
        public object addl_charge_code9 { get; set; }
        public object eta_date { get; set; }
        public object cod_type_text { get; set; }
        public object cod_type { get; set; }
        public object addl_charge_occur3 { get; set; }
        public object reference { get; set; }
        public object sent_to_phone { get; set; }
        public object addl_charge_occur12 { get; set; }
        public object callback_required_text { get; set; }
        public object callback_required { get; set; }
        public object service_level_text { get; set; }
        public object service_level { get; set; }
        public object original_id { get; set; }
        public object width { get; set; }
        public object received_sequence { get; set; }
        public object transfer_to_sequence { get; set; }
        public object cases { get; set; }
        public object times_sent { get; set; }
        public object transfer_to_route { get; set; }
        public object zip_code { get; set; }
        public object settlement_override_amt { get; set; }
        public object driver_app_status_text { get; set; }
        public object driver_app_status { get; set; }
        public object route_code_text { get; set; }
        public object route_code { get; set; }
        public object received_shift { get; set; }
        public object addl_charge_occur6 { get; set; }
        public object addl_charge_occur11 { get; set; }
        public object vehicle { get; set; }
        public object addl_charge_code5 { get; set; }
        public object addl_charge_occur9 { get; set; }
        public object eta { get; set; }
        public object departure_time { get; set; }
        public object combine_data { get; set; }
        public object actual_latitude { get; set; }
        public object posted_by { get; set; }
        public object insurance_value { get; set; }
        public object return_redel_id { get; set; }
        public object addl_charge_code1 { get; set; }
        public object origin_code_text { get; set; }
        public object origin_code { get; set; }
        public object ordered_by { get; set; }
        public object posted_date { get; set; }
        public object actual_billing_amt { get; set; }
        public object created_date { get; set; }
        public object latitude { get; set; }
        public object received_pieces { get; set; }
        public object addl_charge_code7 { get; set; }
        public object totes { get; set; }
        public object asn_sent { get; set; }
        public object comments { get; set; }
        public object verification_id_type_text { get; set; }
        public object verification_id_type { get; set; }
        public object posted_time { get; set; }
        public object item_scans_required { get; set; }
        public object shift_id { get; set; }
        public object addon_billing_amt { get; set; }
        public object actual_delivery_date { get; set; }
        public object id { get; set; }
        public object actual_arrival_time { get; set; }
        public object signature_required { get; set; }
        public object longitude { get; set; }
        public object expected_pieces { get; set; }
        public object loaded_pieces { get; set; }
        public object alt_lookup { get; set; }
        public object customer_number_text { get; set; }
        public object customer_number { get; set; }
        public object created_time { get; set; }
        public object addl_charge_code8 { get; set; }
        public object signature { get; set; }
        public object actual_depart_time { get; set; }
        public object bol_number { get; set; }
        public object actual_cod_type_text { get; set; }
        public object actual_cod_type { get; set; }
        public object invoice_number { get; set; }
        public object branch_id { get; set; }
        public object special_instructions2 { get; set; }
        public object updated_by { get; set; }
        public object verification_id_details { get; set; }
        public object required_signature_type_text { get; set; }
        public object required_signature_type { get; set; }
        public object addl_charge_occur7 { get; set; }
        public object orig_order_number { get; set; }
        public object special_instructions1 { get; set; }
        public List<object> notes { get; set; }
        public object image_sign_req { get; set; }
        public object attention { get; set; }
        public object minutes_late { get; set; }
        public object late_notice_time { get; set; }
        public object received_unique_id { get; set; }
        public object exception_code { get; set; }
        public object addl_charge_code4 { get; set; }
        public object addl_charge_occur4 { get; set; }
        public object redelivery { get; set; }
        public object addl_charge_occur10 { get; set; }
        public object upload_date { get; set; }
        public object special_instructions4 { get; set; }
        public object address_name { get; set; }
        public object addl_charge_occur8 { get; set; }
        public object address_point_customer { get; set; }
        public object received_branch { get; set; }
        public List<object> items { get; set; }
        public object return_redelivery_date { get; set; }
        public object height { get; set; }
        public object actual_longitude { get; set; }
        public object service_time { get; set; }
        public object phone_ext { get; set; }
        public object addl_charge_occur2 { get; set; }
        public object late_notice_date { get; set; }
        public object address { get; set; }
        public object arrival_time { get; set; }
        public object posted_status { get; set; }
        public object route_date { get; set; }
        public object addl_charge_code12 { get; set; }
        public object addl_charge_code3 { get; set; }
        public object return_redelivery_flag_text { get; set; }
        public object return_redelivery_flag { get; set; }
        public object additional_instructions { get; set; }
        public object updated_by_scanner { get; set; }
        public object special_instructions3 { get; set; }
        public object addl_charge_occur5 { get; set; }
        public object address_point { get; set; }
        public object actual_weight { get; set; }
        public object received_company { get; set; }
        public object addl_charge_code2 { get; set; }
        public object state { get; set; }
    }

    public class RouteStopResponseItem
    {
        public object item_number { get; set; }
        public object item_description { get; set; }
        public object reference { get; set; }
        public object rma_route { get; set; }
        public object upload_time { get; set; }
        public object rma_stop_id { get; set; }
        public object width { get; set; }
        public object redelivery { get; set; }
        public object received_pieces { get; set; }
        public object cod_amount { get; set; }
        public object height { get; set; }
        public object comments { get; set; }
        public object actual_pieces { get; set; }
        public object actual_cod_amount { get; set; }
        public object rma_number { get; set; }
        public object manually_updated { get; set; }
        public object unique_id { get; set; }
        public object cod_type_text { get; set; }
        public object cod_type { get; set; }
        public object barcodes_unique { get; set; }
        public object actual_cod_type_text { get; set; }
        public object actual_cod_type { get; set; }
        public object return_redel_seq { get; set; }
        public object expected_pieces { get; set; }
        public object signature { get; set; }
        public object exception_code { get; set; }
        public object company_number_text { get; set; }
        public object company_number { get; set; }
        public object updated_date { get; set; }
        public object expected_weight { get; set; }
        public object created_date { get; set; }
        public object rma_origin { get; set; }
        public object created_by { get; set; }
        public object loaded_pieces { get; set; }
        public object return_redelivery_flag_text { get; set; }
        public object return_redelivery_flag { get; set; }
        public object original_id { get; set; }
        public object container_id { get; set; }
        public object @return { get; set; }
        public object length { get; set; }
        public List<object> notes { get; set; }
        public object actual_weight { get; set; }
        public object updated_by { get; set; }
        public object photos_exist { get; set; }
        public object second_container_id { get; set; }
        public object return_redel_id { get; set; }
        public object asn_sent { get; set; }
        public object actual_departure_time { get; set; }
        public object updated_time { get; set; }
        public object return_redelivery_date { get; set; }
        public object actual_arrival_time { get; set; }
        public object item_sequence { get; set; }
        public object pallet_number { get; set; }
        public object actual_date { get; set; }
        public object insurance_value { get; set; }
        public object created_time { get; set; }
        public object upload_date { get; set; }
        public List<object> scans { get; set; }
        public object id { get; set; }
        public object truck_id { get; set; }
    }


    public class RouteStopInputFile
    {
        public object Customer_Reference { get; set; }
        public object ServiceType { get; set; }
        public object DeliveryName { get; set; }
        public object DeliveryAddress { get; set; }
        public object DeliveryCity { get; set; }
        public object DeliveryState { get; set; }
        public object DeliveryZip { get; set; }
        public object DeliveryPhoneNumber { get; set; }
        public object ItemNumber { get; set; }
        public object ItemDescription { get; set; }
        public object Pieces { get; set; }
        public object Weight { get; set; }
        public object @Return { get; set; }
        public object Bol_Number { get; set; }

    }
    public class Progress
    {
        public object id { get; set; }
        public object status_date { get; set; }
        public object status_text { get; set; }
        public object status_time { get; set; }

    }

    public class RouteStopResponseProgress
    {
        public object status_time { get; set; }
        public object status_date { get; set; }
        public object status_text { get; set; }
        public object id { get; set; }
    }

    public class RouteStopResponseNote
    {
        public object company_number { get; set; }
        public object company_number_text { get; set; }
        public object entry_time { get; set; }
        public object note_text { get; set; }
        public object item_sequence { get; set; }
        public object entry_date { get; set; }
        public object user_entered { get; set; }
        public object show_to_cust { get; set; }
        public object note_type_text { get; set; }
        public object note_type { get; set; }
        public object unique_id { get; set; }
        public object id { get; set; }
        public object user_id { get; set; }
    }
}
