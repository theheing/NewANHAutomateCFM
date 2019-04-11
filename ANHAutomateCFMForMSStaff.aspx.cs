using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using Microsoft.VisualBasic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net.Configuration;
using System.Net.Mail;
using System.Text;
using System.Web.Configuration;
using Automate.Utilities;

using System.Net;
using Microsoft.Reporting.WebForms;

public partial class ANHAutomateCFMForMSStaff : System.Web.UI.Page
    {
        public static string ConnStrLHRSRV = ConfigurationManager.ConnectionStrings["cnLHCSRV"].ToString();
        public static string ConnStrInfSRV = ConfigurationManager.ConnectionStrings["cnInfSRV"].ToString();
        protected void Page_Load(object sender, EventArgs e)
        {
            ServicePointManager.ServerCertificateValidationCallback = Helpers.RemoteCertificateValidationCB;

            //URLTest : ANHAutomateCFMForMSStaff.aspx?status=2&confirmed=1&ReservationID=138711
            int cConfirm = string.IsNullOrEmpty(Request.QueryString["confirmed"]) || (Request.QueryString["confirmed"].ToString() == "0") ? 0 : Convert.ToInt32(Request.QueryString["confirmed"].ToString());
            string cApprovalStatus = string.Empty;
            if (!string.IsNullOrEmpty(Request.QueryString["status"]) && Request.QueryString["status"].ToString() == "2")
                cApprovalStatus = "Approved";
            else if (!string.IsNullOrEmpty(Request.QueryString["status"]) && Request.QueryString["status"].ToString() == "3")
                cApprovalStatus = "Not Approve";
              
            string cReservationID = string.IsNullOrEmpty(Request.QueryString["ReservationID"]) ? "" : Request.QueryString["ReservationID"].ToString(); //Request.QueryString["ReservationID"].ToString();  //"112685";

           if(cReservationID.Length == 0)
                return;
           //string sql = string.Format("SELECT top 1 * FROM  [TSWDATA_ClientCustom].[dbo].[vwANHOnline_Automate_BookingConfirmationLetter] where ReservationID = {0} "
           //                            , cReservationID);
            string messageLog = string.Empty;
            using (SqlConnection cn = new SqlConnection(ConnStrLHRSRV))
            {
                //SqlCommand cmd = new SqlCommand(sql, cn);
                SqlCommand cmd = new SqlCommand("TSWDATA_ClientCustom.laguna.spANHOnline_Automate_BookingConfirmationLetter", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@ReservationID", cReservationID));
                cmd.CommandTimeout = 0;
                cn.Open();

                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string[] recipentEmail = new string[1];
                    recipentEmail[0] = !string.IsNullOrEmpty(reader["StaffEmail"].ToString()) ? reader["StaffEmail"].ToString() : "anvc@club-memberservices.com";
                  
                    string factsheet = string.Empty;
                    for (int i = 0; i < recipentEmail.Length; i++)
                    {


                        string subject = string.Empty;
                        string CXLPolicyCond = string.Empty;
                        string body = string.Empty;

                        //ReservationTypeID	ReservationType
                        //      38	             LLS
                        if(Convert.ToInt32(reader["ReservationTypeID"]) == 38)
                        {
                            StreamReader streader = new StreamReader(Server.MapPath("~/Template/cfm_LLS.html"));
                            string readFile = streader.ReadToEnd();
                            //string body = string.Empty;
                            body = readFile;
                            body = body.Replace("%ContractNumber%", reader["ContractNumber"].ToString());
                            body = body.Replace("%OwnerName%", reader["OwnerName"].ToString());
                            body = body.Replace("%MoreAttacthMsg%", "and Point Statement");
                            body = body.Replace("%SiteName%", reader["SiteName"].ToString());
                            body = body.Replace("%RoomType%", reader["RoomTypeDesc"].ToString());
                            body = body.Replace("%cInDate%", reader["cInDate"].ToString());
                            body = body.Replace("%cOutDate%", reader["cOutDate"].ToString());
                            body = body.Replace("%PointAmt%", reader["PointAmount"] != DBNull.Value ? Convert.ToDouble(reader["PointAmount"].ToString()).ToString("0") : string.Empty);
                            body = body.Replace("%ChkInTime%", reader["CheckInTime"].ToString());
                            body = body.Replace("%ChkOutTime%", reader["CheckOutTime"].ToString());


                            //ReservationSubTypeID	ReservationSubType
                            //340	LLS >=7
                            //341	LLS <7
                            if (reader["ReservationSubType"].ToString() == "LLS >30 days")
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, DateTime.Now.Date, Convert.ToDateTime(reader["InDate"])) > 30)
                                    CXLPolicyCond += "Cancellation or amendment of stay 31 days or more from the first date of the reserved Use period, 100% usage of points returned<br/>" +
                                    "Member will receive 100% usage of points returned if cancellation or amendment of stay within 48 hours of the time the reservation was made. If more than 48 hours, the penalty will be applied as above.";
                                else if (DateAndTime.DateDiff(DateInterval.Day, DateTime.Now.Date, Convert.ToDateTime(reader["InDate"])) <= 30)
                                    CXLPolicyCond += "Member will receive 100% usage of points returned if cancellation or amendment of stay within 48 hours of the time the reservation was made. If more than 48 hours, 100% usage of points forfeited.";

                            }
                            else if (reader["ReservationSubType"].ToString() == "LLS 30 days or less")
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, DateTime.Now.Date, Convert.ToDateTime(reader["InDate"])) > 30)
                                    CXLPolicyCond += "Cancellation or amendment of stay 31 days or more from the first date of the reserved Use period, 100% usage of points returned<br/>" +
                                    "Member will receive 100% usage of points returned if cancellation or amendment of stay within 48 hours of the time the reservation was made. If more than 48 hours, the penalty will be applied as above.";
                                else if (DateAndTime.DateDiff(DateInterval.Day, DateTime.Now.Date, Convert.ToDateTime(reader["InDate"])) <= 30)
                                    CXLPolicyCond += "Member will receive 100% usage of points returned if cancellation or amendment of stay within 48 hours of the time the reservation was made. If more than 48 hours, 100% usage of points forfeited.";

                            }

//                            CXLPolicyCond += @"<p Style='Margin-top:15px;Margin-bottom:15px; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Member will receive 100% usage of points returned if cancellation or amendment of stay within 48 hours of the reservation was made. If more than 48 hours, the penalty will be applied as above.</li>
//                                               </p>
//                                              ";

                            body = body.Replace("%CXLPolicyCond%", CXLPolicyCond);
                            if (cConfirm == 1 && cApprovalStatus == "Approved")
                                subject = reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else if (cApprovalStatus == "Approved")
                                subject = "Approved => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else if (cApprovalStatus == "Not Approve")
                                subject = "Disapproved => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else
                                subject = "Please approve => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();




                        }
                        // LHC MRP Member
                        else if (reader["ContractTypeID"].ToString() == "33" || reader["ContractTypeID"].ToString() == "34" || reader["ContractTypeID"].ToString() == "35")
                        {
                            StreamReader streader = new StreamReader(Server.MapPath("~/Template/cfm_MRP.html"));
                            string readFile = streader.ReadToEnd();
                            //string body = string.Empty;
                            body = readFile;
                            body = body.Replace("%ContractNumber%", reader["ContractNumber"].ToString());
                            body = body.Replace("%OwnerName%", reader["OwnerName"].ToString());
                            body = body.Replace("%MoreAttacthMsg%", "and Point Statement");
                            body = body.Replace("%SiteName%", reader["SiteName"].ToString());
                            body = body.Replace("%RoomType%", reader["RoomTypeDesc"].ToString());
                            body = body.Replace("%cInDate%", reader["cInDate"].ToString());
                            body = body.Replace("%cOutDate%", reader["cOutDate"].ToString());
                            body = body.Replace("%PointAmt%", reader["PointAmount"] != DBNull.Value ? Convert.ToDouble(reader["PointAmount"].ToString()).ToString("0") : string.Empty);
                            body = body.Replace("%ChkInTime%", "15.00 hours");


                            //Old Site
                            //14	LHC @ Twin Peaks Residence
                            //15	LHC @ View Talay Residence 6
                            //16	LHC @ Boathouse Hua Hin

                            //New Site
                            //7	LHC @ Twin Peaks Residence
                            //8	LHC @ View Talay Residence 6
                            //9	LHC @ Boathouse Hua Hin
                            switch (Convert.ToInt32(reader["SiteID"]))
                            {
                                case 7:
                                case 8:
                                case 9:
                                    {
                                        body = body.Replace("%ChkOutTime%", "11.00 hrs.");
                                        break;
                                    }
                                default:
                                    {
                                        body = body.Replace("%ChkOutTime%", "12.00 hrs.");
                                        break;
                                    }
                            }

                            //Old TSW
                            //62	12	AVC 90 or less
                            //66	12	AVC >90
                            //146	12	AVC Plat
                            //126	12	Plat Owner

                            //81    33  LHC 90 or less
                            //87	33	LHC >90
                            //91	33	Plat Owner
                            //147	33	LHC Plat


                            //New TSW
                            //21	AVC 90 or less
                            //17	AVC >90
                            //22	AVC Plat
                            //120	Plat Owner	6 AVC Member

                            //82	LHC 90 or less
                            //78	LHC >90	20 LHC Member
                            //121	Plat Owner	20 LHC Member
                            //85	LHC Plat	20 LHC Member

                            if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 21 ||
                                   Convert.ToInt32(reader["ReservationSubTypeID"]) == 17)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 3)
                                {
                                    CXLPolicyCond += "Short notice confirmed 3-1 days prior to check in date, cancellation or amendment notice must be received within 24 hours of date the booking was confirmed otherwise 100% of points will be deducted.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 14)
                                {
                                    CXLPolicyCond += "Short notice reservation confirmed 14-4 days prior to check-in date , cancellation or amendment notice must be received 3 days prior to requested check in date. Notification less than 3 days notice prior to check in date will result in 100% deduction of points.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) >= 15)
                                {
                                    CXLPolicyCond += "Reservations confirmed  15 days or more in advance, cancellation notices must be received in writing by Lifestyle Services Contact Centre at least  15 days prior to check-in date. Notification less than 15 days prior to check in date will result in 50% deduction of points. Notification less than 3 days prior to check in date will result in 100% deduction of points.";
                                }
                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 22)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 30)
                                {
                                    CXLPolicyCond += "PEAK PERIOD BOOKINGS. Cancellation or amendment notice must be received in writing. Notification less than 30 days prior to check-in date will result in a 50% deduction of points and 100% full forfeit of points will be assessed if amendment or cancellation is received less than 15 days prior to check-in date.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 30)
                                {
                                    CXLPolicyCond += "Short notice for PEAK PERIOD BOOKINGS confirmed 30 days or less prior to check in date, cancellation or amendment must advise LHC within 24 hours of date the booking was confirmed otherwise 100% of points will be deducted.";
                                }
                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 120)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 30)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>100% of points will be returned to your account  for cancellation received in writing by Lifestyle Services Contact Centre 90 days before check in date.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>For cancellation received  89-60 days ,  75% of points  will be returned to your account</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>For cancellation received  59-30 days ,  50%   of points will be returned to your account</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>For cancellation received  30 days  or less  prior to check in, 100% of points will be forfeited </li>
                                              </u>
                                              ";
                                }

                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 82 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 78 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 121 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 85)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 60)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>For reservations confirmed 60 days or less prior to check in , Members may cancel or amendment within 48 hours of the reservation being confirmed by Lifestyle Services Contact Centre with no deduction of points by notifying Lifestyle Services Contact Centre in writing via fax or email.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled more than 30 days prior to check in date will be deducted 50% of points charged for the reservation.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled 29 days or less prior to check-in date will result in 100% of points being deducted.</li>
                                              </u>
                                              ";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 60)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled 60 days or more prior to check-in date will have 100% points credited back to the Member’s account.<br/>
                                                          A Member is entitled to 1 such cancellation per year at no charge; subsequent cancellations will incur an admin charge of 100 points, which will be automatically deducted from the Member’s account.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled 59 – 30 days prior to check-in date will be deducted 50% of points charged for the reservation.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled 29 days or less prior to check-in date will result in 100% of points being deducted.</li>
                                              </u>
                                              ";

                                }


                            }

                            body = body.Replace("%CXLPolicyCond%", CXLPolicyCond);
                            if (cConfirm == 1 && cApprovalStatus == "Approved")
                                subject = reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else if (cApprovalStatus == "Approved")
                                subject = "Approved => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else if (cApprovalStatus == "Not Approve")
                                subject = "Disapproved => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else
                                subject = "Please approve => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();

                        }
                        else if (reader["DefaultLanguage"].ToString() == "EN")
                        {

                            StreamReader streader = new StreamReader(Server.MapPath("~/Template/cfm_en.html"));
                            string readFile = streader.ReadToEnd();
                            //string body = string.Empty;
                            body = readFile;
                            body = body.Replace("%ContractNumber%", reader["ContractNumber"].ToString());
                            body = body.Replace("%OwnerName%", reader["OwnerName"].ToString());
                            body = body.Replace("%MoreAttacthMsg%", "and Point Statement" );
                            body = body.Replace("%SiteName%", reader["SiteName"].ToString());
                            body = body.Replace("%RoomType%", reader["RoomTypeDesc"].ToString());
                            body = body.Replace("%cInDate%", reader["cInDate"].ToString());
                            body = body.Replace("%cOutDate%", reader["cOutDate"].ToString());
                            body = body.Replace("%PointAmt%", reader["PointAmount"] != DBNull.Value ? Convert.ToDouble(reader["PointAmount"].ToString()).ToString("0") : string.Empty);
                            body = body.Replace("%ChkInTime%", "15.00 hours");


                            //Old Site
                            //14	LHC @ Twin Peaks Residence
                            //15	LHC @ View Talay Residence 6
                            //16	LHC @ Boathouse Hua Hin

                            //New Site
                            //7	LHC @ Twin Peaks Residence
                            //8	LHC @ View Talay Residence 6
                            //9	LHC @ Boathouse Hua Hin
                            switch (Convert.ToInt32(reader["SiteID"]))
                            {
                                case 7:
                                case 8:
                                case 9:
                                    {
                                        body = body.Replace("%ChkOutTime%", "11.00 hrs.");
                                        break;
                                    }
                                default:
                                    {
                                        body = body.Replace("%ChkOutTime%", "12.00 hrs.");
                                        break;
                                    }
                            }

                            //Old TSW
                            //62	12	AVC 90 or less
                            //66	12	AVC >90
                            //146	12	AVC Plat
                            //126	12	Plat Owner

                            //81    33  LHC 90 or less
                            //87	33	LHC >90
                            //91	33	Plat Owner
                            //147	33	LHC Plat


                            //New TSW
                            //21	AVC 90 or less
                            //17	AVC >90
                            //22	AVC Plat
                            //120	Plat Owner	6 AVC Member

                            //82	LHC 90 or less
                            //78	LHC >90	20 LHC Member
                            //121	Plat Owner	20 LHC Member
                            //85	LHC Plat	20 LHC Member
                            
                            if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 21 ||
                                   Convert.ToInt32(reader["ReservationSubTypeID"]) == 17)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 3)
                                {
                                    CXLPolicyCond += "Short notice confirmed 3-1 days prior to check in date, cancellation or amendment notice must be received within 24 hours of date the booking was confirmed otherwise 100% of points will be deducted.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 14)
                                {
                                    CXLPolicyCond += "Short notice reservation confirmed 14-4 days prior to check-in date , cancellation or amendment notice must be received 3 days prior to requested check in date. Notification less than 3 days notice prior to check in date will result in 100% deduction of points.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) >= 15)
                                {
                                    CXLPolicyCond += "Reservations confirmed  15 days or more in advance, cancellation notices must be received in writing by Club Member Services at least  15 days prior to check-in date. Notification less than 15 days prior to check in date will result in 50% deduction of points. Notification less than 3 days prior to check in date will result in 100% deduction of points.";
                                }
                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 22)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 30)
                                {
                                    CXLPolicyCond += "PEAK PERIOD BOOKINGS. Cancellation or amendment notice must be received in writing. Notification less than 30 days prior to check-in date will result in a 50% deduction of points and 100% full forfeit of points will be assessed if amendment or cancellation is received less than 15 days prior to check-in date.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 30)
                                {
                                    CXLPolicyCond += "Short notice for PEAK PERIOD BOOKINGS confirmed 30 days or less prior to check in date, cancellation or amendment must advise LHC within 24 hours of date the booking was confirmed otherwise 100% of points will be deducted.";
                                }
                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 120)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 30)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>100% of points will be returned to your account  for cancellation received in writing by Club Member Services 90 days before check in date.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>For cancellation received  89-60 days ,  75% of points  will be returned to your account</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>For cancellation received  59-30 days ,  50%   of points will be returned to your account</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>For cancellation received  30 days  or less  prior to check in, 100% of points will be forfeited </li>
                                              </u>
                                              ";
                                }

                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 82 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 78 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 121 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 85)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 60)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>For reservations confirmed 60 days or less prior to check in , Members may cancel or amendment within 48 hours of the reservation being confirmed by Club Member Services with no deduction of points by notifying Club Member Services  in writing via fax or email.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled more than 30 days prior to check in date will be deducted 50% of points charged for the reservation.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled 29 days or less prior to check-in date will result in 100% of points being deducted.</li>
                                              </u>
                                              ";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 60)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled 60 days or more prior to check-in date will have 100% points credited back to the Member’s account.<br/>
                                                          A Member is entitled to 1 such cancellation per year at no charge; subsequent cancellations will incur an admin charge of 100 points, which will be automatically deducted from the Member’s account.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled 59 – 30 days prior to check-in date will be deducted 50% of points charged for the reservation.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;Margin-left: 15px;font-size:11.0pt;font-family:&quot;Arial Narrow&quot;,&quot;sans-serif&quot;;'>Reservations cancelled 29 days or less prior to check-in date will result in 100% of points being deducted.</li>
                                              </u>
                                              ";

                                }


                            }

                            body = body.Replace("%CXLPolicyCond%", CXLPolicyCond);
                            if (cConfirm == 1 && cApprovalStatus == "Approved")
                                subject =  reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else if(cApprovalStatus == "Approved")
                                subject = "Approved => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else if(cApprovalStatus == "Not Approve")
                                subject = "Disapproved => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();
                            else
                                subject = "Please approve => " + reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();


                        }
                        else if (reader["DefaultLanguage"].ToString() == "TH")
                        {

                            StreamReader streader = new StreamReader(Server.MapPath("~/Template/cfm_th.html"));
                            string readFile = streader.ReadToEnd();
                            //string body = string.Empty;
                            body = readFile;
                            body = body.Replace("%ContractNumber%", reader["ContractNumber"].ToString());
                            body = body.Replace("%OwnerName%", reader["OwnerName"].ToString());
                            body = body.Replace("%MoreAttacthMsg%",  "และ เอกสารแสดงคะแนน" );
                            body = body.Replace("%SiteName%", reader["SiteName"].ToString());
                            body = body.Replace("%RoomType%", reader["RoomTypeDesc"].ToString());
                            body = body.Replace("%cInDate%", reader["cInDate"].ToString());
                            body = body.Replace("%cOutDate%", reader["cOutDate"].ToString());
                            body = body.Replace("%PointAmt%", reader["PointAmount"] != DBNull.Value ? Convert.ToDouble(reader["PointAmount"].ToString()).ToString("0") : string.Empty);
                            body = body.Replace("%ChkInTime%", "15.00 น.");
                            //Old Site
                            //14	LHC @ Twin Peaks Residence
                            //15	LHC @ View Talay Residence 6
                            //16	LHC @ Boathouse Hua Hin

                            //New Site
                            //7	LHC @ Twin Peaks Residence
                            //8	LHC @ View Talay Residence 6
                            //9	LHC @ Boathouse Hua Hin
                            switch (Convert.ToInt32(reader["SiteID"]))
                            {
                                case 7:
                                case 8:
                                case 9:
                                    {
                                        body = body.Replace("%ChkOutTime%", "11.00 น.");
                                        break;
                                    }
                                default:
                                    {
                                        body = body.Replace("%ChkOutTime%", "12.00 น.");
                                        break;
                                    }
                            }


                            //Old TSW
                            //62	12	AVC 90 or less
                            //66	12	AVC >90
                            //146	12	AVC Plat
                            //126	12	Plat Owner

                            //81    33  LHC 90 or less
                            //87	33	LHC >90
                            //91	33	Plat Owner
                            //147	33	LHC Plat


                            //New TSW
                            //21	AVC 90 or less
                            //17	AVC >90
                            //22	AVC Plat
                            //120	Plat Owner	6 AVC Member

                            //82	LHC 90 or less
                            //78	LHC >90	20 LHC Member
                            //121	Plat Owner	20 LHC Member
                            //85	LHC Plat	20 LHC Member

                            if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 21 ||
                                   Convert.ToInt32(reader["ReservationSubTypeID"]) == 17)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 3)
                                {
                                    CXLPolicyCond += "สำหรับคำขอจองห้องพักที่ได้รับการยืนยัน ถูกร้องขอในช่วงเวลาอันจำกัด ภายใน 3-1 วันก่อนวันลงทะเบียนเข้าพัก การยกเลิก หรือเปลี่ยนแปลงการจองห้องพักต้องแจ้งให้ ทีมบริการสมาชิกทราบเป็นลายลักษณ์อักษรภายใน 24 ชั่วโมงนับจากวันที่ได้รับการยืนยันการเข้าพัก  มิฉะนั้นจะถูกหักคะแนน 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 14)
                                {
                                    CXLPolicyCond += "สำหรับคำขอจองห้องพักที่ได้รับการยืนยัน ถูกร้องขอในช่วงเวลาอันจำกัด ภายใน 14-4 วัน ก่อนวันลงทะเบียนเข้าพัก การยกเลิก หรือเปลี่ยนแปลงการจองห้องพักต้องแจ้งให้ทีมบริการสมาชิก ทราบเป็นลายลักษณ์อักษรมากกว่า 3 วันล่วงหน้าก่อนวันลงทะเบียนเข้าพัก  หากแจ้งน้อยกว่า 3 วัน ก่อนวันลงทะเบียนเข้าพัก จะถูกหักคะแนน 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) >= 15)
                                {
                                    CXLPolicyCond += "สำหรับคำขอจองห้องพักที่ได้รับการยืนยัน มากกว่า 15 วัน ก่อนวันลงทะเบียนเข้าพัก การยกเลิก หรือเปลี่ยนแปลงการจองห้องพักต้องแจ้งให้ทีมบริการสมาชิกทราบเป็นลายลักษณ์อักษรมากกว่า 15 วันล่วงหน้าก่อนวันลงทะเบียนเข้าพัก หากแจ้งน้อยกว่า 15 วัน จะถูกหักคะแนน 50 เปอร์เซ็นต์ ของคะแนนที่ใช้ในการสำรองห้องพัก. และจะถูกหักคะแนน 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก หากแจ้งน้อยกว่า 3 วัน ก่อนวันลงทะเบียนเข้าพัก ";
                                }
                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 22)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 30)
                                {
                                    CXLPolicyCond += "สำหรับคำขอจองห้องพักที่ได้รับการยืนยัน ถูกร้องขอในช่วงเวลาอันจำกัดและเข้าสู่ช่วงเทศกาล ภายใน 30 วันหรือน้อยกว่าก่อนวันลงทะเบียนเข้าพัก การยกเลิกหรือเปลี่ยนแปลงการจองห้องพักต้องแจ้งให้ทีมบริการสมาชิกทราบเป็นลายลักษณ์อักษรภายใน 24 ชั่วโมงนับจากวันที่ได้รับการยืนยันการเข้าพัก  มิฉะนั้นจะถูกหักคะแนน 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 30)
                                {
                                    CXLPolicyCond += "สำหรับคำขอจองห้องพักที่ได้รับการยืนยันเข้าสู่ช่วงเทศกาล  การยกเลิกหรือเปลี่ยนแปลงการจองห้องพักต้องแจ้งให้ทีมบริการสมาชิก ทราบเป็นลายลักษณ์อักษรมากกว่า 30 วันล่วงหน้าก่อนวันลงทะเบียนเข้าพัก หากแจ้งน้อยกว่า 30 วัน จะถูกหักคะแนน 50 เปอร์เซ็นต์ และจะถูกหัก 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก หากแจ้งน้อยกว่า 15 วัน ก่อนวันลงทะเบียนเข้าพัก";
                                }
                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 120)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 30)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>การยกเลิก หรือเปลีี่ยนแปลงการจองห้องพักที่แจ้งให้ทีมบริการสมาชิกทราบเป็นลายลักษณ์อักษรมากกว่า 90 วันล่วงหน้าก่อนวันลงทะเบียนเข้าพัก ท่านจะได้รับ คะแนน คืนกลับสู่บัญชี 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก.</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>การยกเลิก หรือเปลี่ยนแปลงการจองห้องพักที่แจ้งให้ทีมบริการสมาชิก ทราบเป็นลายลักษณ์อักษรภายใน 89-60วันล่วงหน้าก่อนวันลงทะเบียนเข้าพัก ท่านจะได้รับ คะแนน คืนกลับสู่บัญชี 75 เปอร์เซ็นต์</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>การยกเลิก หรือเปลี่ยนแปลงการจองห้องพักที่แจ้งให้ทีมบริการสมาชิก ทราบเป็นลายลักษณ์อักษรภายใน 59-30 วันล่วงหน้าก่อนวันลงทะเบียนเข้าพัก ท่านจะได้รับ คะแนน คืนกลับสู่บัญชี 50 เปอร์เซ็นต์</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>การยกเลิก หรือเปลี่ยนแปลงการจองห้องพักที่แจ้งให้ทีมบริการสมาชิก ทราบเป็นลายลักษณ์อักษรน้อยกว่า  30 ก่อนวันลงทะเบียนเข้าพัก ท่านจะถูกหักคะแนน 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก.</li>
                                              </u>
                                              ";
                                }

                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 82 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 78 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 121 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 85)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 60)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>สำหรับคำขอจองห้องพักที่ได้รับการยืนยัน 60 วันหรือน้อยกว่าก่อนวันลงทะเบียนเข้าพัก ( check in date ) สมาชิกสามารถยกเลิกหรือเปลี่ยนแปลง คำยืนยันการจองห้องพักได้ภายใน 48 ชั่วโมงนับจากวันที่ได้รับการยืนยันการจองห้องพักจากทีมบริการสมาชิก   จะไม่ถูกหักคะแนนสิทธิ  โดยแจ้งให้ทีมบริการสมาชิก ทราบเป็นลายลักษณ์อักษรผ่านทางโทรสารหรืออีเมล์ </li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>หากมีการยกเลิกการสำรองห้องพักมากกว่า 30 วันก่อนวันลงทะเบียนเข้าพัก (check in date) คะแนนสิทธิจำนวนเท่ากับร้อยละ 50 ของคะแนนสิทธิที่ใช้ไปในการจองห้องพักนั้นๆ เท่านั้น จะคืนกลับเข้าสู่บัญชีคะแนนสิทธิของสมาชิก</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>หากมีการยกเลิกในระหว่าง  29 วันหรือน้อยกว่านั้นก่อนวันลงทะเบียนเข้าพัก (check-in date) จะเป็นทำผลให้คะแนนสิทธิทั้งหมดร้อยละ 100 ของคะแนนสิทธิที่ใช้ไปในการจองห้องพักนั้น ๆ ไม่ได้คืนกลับเข้าสู่บัญชีคะแนนสิทธิของสมาชิก</li>
                                              </u>
                                              ";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 60)
                                {
                                    CXLPolicyCond += @"<u style='text-decoration: none;Margin-top:0;Margin-bottom:0;'>
                                                  <li Style='Margin-top:15px;Margin-bottom:0; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>คำขอจองห้องพักที่ได้รับการยืนยันแล้ว หากมีการยกเลิกภายใน  60 วันหรือมากกว่านั้นก่อนวันลงทะเบียนเข้าพัก (check-in date) จะได้รับคะแนนสิทธิที่หักทั้งหมดคืนกลับเข้าสู่บัญชีคะแนนสิทธิของสมาชิก โดยสมาชิกแต่ละรายมีสิทธิดำเนินการยกเลิกในลักษณะดังกล่าวได้ปีละหนึ่งครั้งโดยจะไม่มีการคิดค่าใช้จ่ายใด ๆ ทั้งสิ้น การยกเลิกการจองหลังจากนั้นในปีเดียวกันจะถือว่ามีค่าธรรมเนียมการบริหารจัดการ (admin charge) ในจำนวนเท่ากับคะแนนสิทธิ 100 คะแนนซึ่งจะหักออกจากบัญชีคะแนนสิทธิของสมาชิกโดยอัตโนมัติ</li>
                                                  <li Style='Margin-top:0;Margin-bottom:0; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>คำขอจองห้องพักที่ได้รับการยืนยันแล้ว หากมีการยกเลิกในระหว่าง  59 - 30 วันก่อนวันลงทะเบียนเข้าพัก (check-in date) คะแนนสิทธิจำนวนเท่ากับร้อยละ  50 ของคะแนนสิทธิที่ใช้ไปในการจองห้องพักนั้นๆ เท่านั้น จะคืนกลับเข้าสู่บัญชีคะแนนสิทธิของสมาชิก</li>
                                                  <li Style='Margin-top:0;Margin-bottom:15px; line-height:18px !important;font-size:10.0pt;font-family:Tahoma, Geneva, sans-serif;Margin-left: 15px;'>คำขอจองห้องพักที่ได้รับการยืนยันแล้ว หากมีการยกเลิกในระหว่าง  29 วันหรือน้อยกว่านั้นก่อนวันลงทะเบียนเข้าพัก (check-in date) จะเป็นทำผลให้คะแนนสิทธิทั้งหมดร้อยละ  100 ของคะแนนสิทธิที่ใช้ไปในการจองห้องพักนั้น ๆ ไม่ได้คืนกลับเข้าสู่บัญชีคะแนนสิทธิของสมาชิก</li>
                                              </u>
                                              ";

                                }


                            }
                           
                           
                            body = body.Replace("%CXLPolicyCond%", CXLPolicyCond);
                            if (cConfirm == 1 && cApprovalStatus == "Approved")
                                subject = reader["CIDMemberNumber"].ToString() + " : จดหมายยืนยันการสำรองห้องพัก - " + reader["SiteName"].ToString();
                            else if (cApprovalStatus == "Approved")
                                subject = "Approved => "+reader["CIDMemberNumber"].ToString()+" : จดหมายยืนยันการสำรองห้องพัก - " + reader["SiteName"].ToString();
                            else if (cApprovalStatus == "Not Approve")
                                subject = "Disapproved => " + reader["CIDMemberNumber"].ToString() + " : จดหมายยืนยันการสำรองห้องพัก - " + reader["SiteName"].ToString();
                            else
                                subject = "Please approve => " + reader["CIDMemberNumber"].ToString() + " : จดหมายยืนยันการสำรองห้องพัก - " + reader["SiteName"].ToString();
                        }






                        body = body.Replace("%Tracking%", string.Empty);


                        MailMessage Msg = new MailMessage();
                        Msg.Subject = subject;
                        Msg.BodyEncoding = Encoding.GetEncoding("UTF-8");
                        Msg.SubjectEncoding = Encoding.GetEncoding("UTF-8");
                        Msg.From = new MailAddress("anvc@club-memberservices.com", "No reply");
                        Msg.ReplyToList.Add(new MailAddress("anvc@club-memberservices.com", "Angsana Vacation Club"));
                        if (cConfirm == 1 && cApprovalStatus == "Approved")
                        {
                            Msg.From = new MailAddress("anvc@club-memberservices.com", "Angsana Vacation Club");
                            Msg.To.Add(new MailAddress("anvc@club-memberservices.com", "Angsana Vacation Club"));
                            Msg.CC.Add(new MailAddress(recipentEmail[i], recipentEmail[i]));
                        }
                        else if (cApprovalStatus == "Approved")
                        {
                           
                            Msg.To.Add(new MailAddress(recipentEmail[i], recipentEmail[i]));

                        }
                        else
                        {

                            Msg.From = new MailAddress("anvc@club-memberservices.com", recipentEmail[i]);
                            if (!string.IsNullOrEmpty(reader["ApproverEmail"].ToString()))
                            {
                               
                                Msg.To.Add(new MailAddress(reader["ApproverEmail"].ToString(), reader["ApproverEmail"].ToString()));
                                Msg.CC.Add(new MailAddress(recipentEmail[i], recipentEmail[i]));
                            }
                            else
                            {
                               Msg.To.Add(new MailAddress(recipentEmail[i], recipentEmail[i]));
                            }
                           

                        }
                       
                        //Msg.Bcc.Add(new MailAddress("eakkaphols@lagunaphuket.com", "eakkaphols@lagunaphuket.com"));
                        Msg.Bcc.Add(new MailAddress("onlineemailtracking@lagunaphuket.com", "onlineemailtracking@lagunaphuket.com"));
                       
                        Msg.Body = body.ToString();
                        Msg.IsBodyHtml = true;

                        if (Convert.ToInt32(reader["ReservationTypeID"]) == 38)
                        {
                            //var contentID = "logo";
                            var lagunaLogo = new Attachment(Server.MapPath("~/Template/images/LagunaLifeStyle_logo.png"));
                            lagunaLogo.ContentId = "LagunaLifeStylelogo";
                            lagunaLogo.ContentDisposition.Inline = true;
                            lagunaLogo.ContentDisposition.DispositionType = System.Net.Mime.DispositionTypeNames.Inline;
                            Msg.Attachments.Add(lagunaLogo);
                        }
                        // LHC MRP Member
                        else if (reader["ContractTypeID"].ToString() == "33" || reader["ContractTypeID"].ToString() == "34" || reader["ContractTypeID"].ToString() == "35")
                        {
                            //var contentID = "logo";
                            var ANHLogo = new Attachment(Server.MapPath("~/Template/images/ANVC_logo_small.jpg"));
                            ANHLogo.ContentId = "ANHlogo";
                            ANHLogo.ContentDisposition.Inline = true;
                            ANHLogo.ContentDisposition.DispositionType = System.Net.Mime.DispositionTypeNames.Inline;
                            Msg.Attachments.Add(ANHLogo);
                        }
                        else
                        {
                            //var contentID = "logo";
                            var ANHLogo = new Attachment(Server.MapPath("~/Template/images/ANH_logo.png"));
                            ANHLogo.ContentId = "ANHlogo";
                            ANHLogo.ContentDisposition.Inline = true;
                            ANHLogo.ContentDisposition.DispositionType = System.Net.Mime.DispositionTypeNames.Inline;
                            Msg.Attachments.Add(ANHLogo);
                        }


                        //var lifeStyleLogo = new Attachment(Server.MapPath("~/Template/images/LifeStyle_logo.png"));
                        //lifeStyleLogo.ContentId = "LifeStyle_logo";
                        //lifeStyleLogo.ContentDisposition.Inline = true;
                        //lifeStyleLogo.ContentDisposition.DispositionType = System.Net.Mime.DispositionTypeNames.Inline;
                        //Msg.Attachments.Add(lifeStyleLogo);





                        /* SSRS Report */
                        Warning[] warnings;
                        string[] streamIds;
                        string mimeType;
                        string encoding;
                        string extension;


                        

                        ReportViewer rv = new ReportViewer { ProcessingMode = ProcessingMode.Remote };
                        IReportServerCredentials irsc = new CustomReportCredentials(Helpers.ReportUserName, Helpers.ReportPassword, "laguna");
                        rv.ServerReport.ReportServerCredentials = irsc;
                        rv.AsyncRendering = false;
                        ServerReport sr = rv.ServerReport;
                        sr.ReportServerUrl = new Uri(Helpers.ReportServer);
 

                        /*Add report  Booking Confirmation Letter*/
                        //33	LHC MRP Convert Ult	
                        //34	LHC MRP Convert Plus	
                        //35	LHC MRP Convert Core	
                        


                        Microsoft.Reporting.WebForms.ReportParameter[] parameters;
                        if (Convert.ToInt32(reader["ReservationTypeID"]) == 38)
                        {
                            sr.ReportPath = "/TSW Custom Reports/LHC Reservation/LHC Member Booking Confirmation (Laguna Lifestyle)";
                            parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                            parameters[0] = new ReportParameter("ResortName", reader["SiteName"].ToString(), true);
                            parameters[1] = new ReportParameter("ConfirmationNumber", reader["ReservationNumber"].ToString(), true);
                            parameters[2] = new ReportParameter("ShowRoomNo", "true", true);
                            
                            rv.ServerReport.SetParameters(parameters);
                        }
                        //MRP
                        else if (reader["ContractTypeID"].ToString() == "33" || reader["ContractTypeID"].ToString() == "34" || reader["ContractTypeID"].ToString() == "35")
                            switch (Convert.ToInt32(reader["SiteID"]))
                            {   case 2 : case 3 : case 4 : case 5 : case 6 : case 7: case 8 : case 9 : case 16 : case 17 : case 19 : case 21 : case 22 : case 52 : case 143 : case 144 : case 145 : case 299 : case 300:
                                    sr.ReportPath = "/TSW Custom Reports/LHC Member Services/LHC Booking Confirmation Letter (English) for LHC MRP";
                                    parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                                    parameters[0] = new ReportParameter("ResortName", reader["SiteName"].ToString(), true);
                                    parameters[1] = new ReportParameter("ReservationNumber", reader["ReservationNumber"].ToString(), true);
                                    parameters[2] = new ReportParameter("ShowRoomNo", "true", true);
                                    rv.ServerReport.SetParameters(parameters);
                                    break;
                                default:
                                    sr.ReportPath = "/TSW Custom Reports/LHC Member Services/ANVC Booking Confirmation Letter (English) for LHC MRP";
                                    parameters = new Microsoft.Reporting.WebForms.ReportParameter[2];
                                    parameters[0] = new ReportParameter("ResortName", reader["SiteName"].ToString(), true);
                                    parameters[1] = new ReportParameter("ConfirmationNumber", reader["ReservationNumber"].ToString(), true);
                                    rv.ServerReport.SetParameters(parameters);
                                    break;
                            }
                        else if (reader["DefaultLanguage"].ToString() == "EN")
                        {
                            
                            sr.ReportPath = "/TSW Custom Reports/LHC Reservation/LHC Member Booking Confirmation (English)";
                            parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                            parameters[0] = new ReportParameter("ResortName", reader["SiteName"].ToString(), true);
                            parameters[1] = new ReportParameter("ReservationNumber", reader["ReservationNumber"].ToString(), true);
                            parameters[2] = new ReportParameter("ShowRoomNo", "true", true);
                            rv.ServerReport.SetParameters(parameters);
                        }
                        else if (reader["DefaultLanguage"].ToString() == "TH")
                        {
                            sr.ReportPath = "/TSW Custom Reports/LHC Reservation/LHC Member Booking Confirmation (Thai)";
                            parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                            parameters[0] = new ReportParameter("ResortName", reader["SiteName"].ToString(), true);
                            parameters[1] = new ReportParameter("ReservationNumber", reader["ReservationNumber"].ToString(), true);
                            parameters[2] = new ReportParameter("ShowRoomNo", "true", true);
                            rv.ServerReport.SetParameters(parameters);
                            
                        }

                        // pass crendentitilas
                       // rptViewer.ServerReport.ReportServerCredentials =  new ReportServerCredentials("uName", "PassWORD", "doMain");
                        //rv.ServerReport.ReportServerCredentials = CustomReportCredentials.//System.Net.CredentialCache.DefaultCredentials  //new ReportServerCredentials("laguna\adselfservice", "Laguna#2018", "laguna");

                        byte[] bytes = rv.ServerReport.Render("PDF", null, out mimeType, out encoding, out extension, out streamIds, out warnings);
                        MemoryStream ms = new MemoryStream(bytes);
                        Msg.Attachments.Add(new Attachment(ms, reader["CIDMemberNumber"].ToString() + "-CFMLetter.pdf"));



                        /*Add report PointStatement*/
                        if (Convert.ToInt32(reader["ReservationTypeID"]) == 38)
                        {
                            sr.ReportPath = "/TSW Custom Reports/LHC Reservation/LHC Point Statement (Laguna Lifestyle)";
                            parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                            parameters[0] = new ReportParameter("ContractNumber", reader["CIDMemberNumber"].ToString(), true);
                            parameters[1] = new ReportParameter("ShowTransactionSummary", "0", true);
                            parameters[2] = new ReportParameter("ShowExpirePoint", "0", true);
                            rv.ServerReport.SetParameters(parameters);
                        }
                        //MRP
                        else if (reader["ContractTypeID"].ToString() == "33" || reader["ContractTypeID"].ToString() == "34" || reader["ContractTypeID"].ToString() == "35")
                        {
                            sr.ReportPath = "/TSW Custom Reports/LHC Member Services/LHC Point Statement (Active Contract) for MRP";
                            parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                            parameters[0] = new ReportParameter("ContractNumber", reader["CIDMemberNumber"].ToString(), true);
                            parameters[1] = new ReportParameter("ShowTransactionSummary", "0", true);
                            parameters[2] = new ReportParameter("ShowExpirePoint", "0", true);
                            rv.ServerReport.SetParameters(parameters);


                        }
                        else if (reader["ReservationType"].ToString() == "LHC Member" && reader["ContractNumber"].ToString().Substring(0, 4) == "ANVC")
                        {
                            sr.ReportPath = "/TSW Custom Reports/ANVC Reservation/ANVC Point Statement (Active Contract)";
                            parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                            parameters[0] = new ReportParameter("Number", reader["CIDMemberNumber"].ToString(), true);
                            parameters[1] = new ReportParameter("ShowTransactionSummary", "True", true);
                            parameters[2] = new ReportParameter("ShowExpirePoint", "false", true);
                            rv.ServerReport.SetParameters(parameters);

                        }
                        else if (reader["DefaultLanguage"].ToString() == "EN")
                        {
                            sr.ReportPath = "/TSW Custom Reports/LHC Reservation/LHC Point Statement (Active Contract)";
                            parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                            parameters[0] = new ReportParameter("ContractNumber", reader["CIDMemberNumber"].ToString(), true);
                            parameters[1] = new ReportParameter("ShowTransactionSummary", "0", true);
                            parameters[2] = new ReportParameter("ShowExpirePoint", "0", true);
                            rv.ServerReport.SetParameters(parameters);

                        }
                        else if (reader["DefaultLanguage"].ToString() == "TH")
                        {
                            sr.ReportPath = "/TSW Custom Reports/LHC Reservation/LHC Point Statement (Active Contract-Thai)";
                            parameters = new Microsoft.Reporting.WebForms.ReportParameter[3];
                            parameters[0] = new ReportParameter("ContractNumber", reader["CIDMemberNumber"].ToString(), true);
                            parameters[1] = new ReportParameter("ShowTransactionSummary", "0", true);
                            parameters[2] = new ReportParameter("ShowExpirePoint", "0", true);
                            rv.ServerReport.SetParameters(parameters);

                        }


                        //rv.ServerReport.ReportServerCredentials = new ReportServerCredentials("eakkaphols", "Laguna#2016", "laguna");
                        bytes = rv.ServerReport.Render("PDF", null, out mimeType, out encoding, out extension, out streamIds, out warnings);
                        ms = new MemoryStream(bytes);
                        Msg.Attachments.Add(new Attachment(ms, reader["CIDMemberNumber"].ToString() + "-PointStatement.pdf"));





                        if (Convert.ToInt32(reader["ReservationTypeID"]) != 38)
                        {
                            /* New TSW Site
                            2	LHC @ Allamanda Laguna Phuket
                            3	LHC @ Angsana Laguna Phuket
                            4	LHC @ Angsana Resort & Spa Bintan
                            5	Laguna Holiday Club Phuket Resort
                            7	LHC @ Twin Peaks Residence
                            8	LHC @ View Talay Residence 6
                            9	LHC @ Boathouse Hua Hin
                            19	LHC Private Pool Villas
                            */
                            int siteID = Convert.ToInt32(reader["SiteID"]);
                            switch (Convert.ToInt32(reader["SiteID"]))
                            {
                                case 2:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["ALMD"]);
                                    break;
                                case 3:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["ANLP"]);
                                    break;
                                case 4:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["ANBT"]);
                                    break;
                                case 5:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["LHCR"]);
                                    break;
                                case 7:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["TWP"]);
                                    break;
                                case 8:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["VTR"]);
                                    break;
                                case 9:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["BHH"]);
                                    break;
                                case 19:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["PPV"]);
                                    break;
                                default:
                                    break;
                            }
                            if (!string.IsNullOrEmpty(factsheet))
                            {
                                Attachment attachment = new Attachment(factsheet);
                                attachment.Name = "Factsheet.pdf";
                                Msg.Attachments.Add(attachment);
                            }
                        }
                        
                        

                        




                        messageLog += string.Format("|   {0}   |   {1}   |   {2}   |\r\n", reader["OwnerEmail"].ToString()
                                                                                         , Helpers.SendHTMLMail(Msg) == true ? "Pass" : "False"
                                                                                         , DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        messageLog += string.Format("----------------------------------------------------------------------\r\n");

                        Msg = null;





                    } //end for

                }
                    reader.Dispose();



                    
                
                
            }

        }



        
    }

public class CustomReportCredentials : IReportServerCredentials
{
    private string _UserName;
    private string _PassWord;
    private string _DomainName;

    public CustomReportCredentials(string UserName, string PassWord, string DomainName)
    {
        _UserName = UserName;
        _PassWord = PassWord;
        _DomainName = DomainName;
    }

    public System.Security.Principal.WindowsIdentity ImpersonationUser
    {
        get { return null; }
    }

    public ICredentials NetworkCredentials
    {
        get { return new NetworkCredential(_UserName, _PassWord, _DomainName); }
    }

    public bool GetFormsCredentials(out Cookie authCookie, out string user,
     out string password, out string authority)
    {
        authCookie = null;
        user = password = authority = null;
        return false;
    }
}

