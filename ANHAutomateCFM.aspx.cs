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
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using Automate.Utilities;



public partial class ANHAutomateCFM : System.Web.UI.Page
    {
        public static string ConnStrLHRSRV = ConfigurationManager.ConnectionStrings["cnLHCSRV"].ToString();
        public static string ConnStrInfSRV = ConfigurationManager.ConnectionStrings["cnInfSRV"].ToString();
        protected void Page_Load(object sender, EventArgs e)
        {

            string sql = string.Format("SELECT * FROM  [TSWDATA_ClientCustom].[dbo].[vwANHOnline_Automate_BookingConfirmationLetter] where  Approval = 'Approved' and PrintConfirmation = 0 and PrintDate = DATEADD(day, DATEDIFF(day, 0, GETDATE()), -1)");
            string messageLog = string.Empty;
            using (SqlConnection cn = new SqlConnection(ConnStrLHRSRV))
            {
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.CommandType = CommandType.Text;
                cmd.CommandTimeout = 0;
                cn.Open();

                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string[] recipentEmail = new string[1];
                    recipentEmail[0] = reader["StaffEmail"].ToString(); //"eakkaphols@lagunaphuket.com"; // reader["StaffEmail"].ToString();
                    string factsheet = string.Empty;
                    for (int i = 0; i < recipentEmail.Length; i++)
                    {


                        string subject = string.Empty;
                        string CXLPolicyCond = string.Empty;
                        string body = string.Empty;
                        if (reader["DefaultLanguage"].ToString() == "EN")
                        {

                            StreamReader streader = new StreamReader(Server.MapPath("~/Template/cfm_en.html"));
                            string readFile = streader.ReadToEnd();
                            //string body = string.Empty;
                            body = readFile;
                            body = body.Replace("%ContractNumber%", reader["ContractNumber"].ToString());
                            body = body.Replace("%OwnerName%", reader["OwnerName"].ToString());
                            body = body.Replace("%MoreAttacthMsg%","and Point Statement");
                            body = body.Replace("%SiteName%", reader["SiteName"].ToString());
                            body = body.Replace("%RoomType%", reader["RoomTypeDesc"].ToString());
                            body = body.Replace("%cInDate%", reader["cInDate"].ToString());
                            body = body.Replace("%cOutDate%", reader["cOutDate"].ToString());
                            body = body.Replace("%PointAmt%", reader["PointAmount"] != DBNull.Value ? Convert.ToDouble(reader["PointAmount"].ToString()).ToString("0") : string.Empty);
                            body = body.Replace("%ChkInTime%", "15.00 hours");
                            switch (Convert.ToInt32(reader["SiteID"]))
                            {
                                case 14:
                                case 15:
                                case 16:
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


                            if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 62 ||
                                   Convert.ToInt32(reader["ReservationSubTypeID"]) == 66)
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
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 146)
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
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 126)
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
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 81 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 87 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 91 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 147)
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
                            subject = reader["CIDMemberNumber"].ToString() + " : Confirmation Letter - " + reader["SiteName"].ToString();


                        }
                        else if (reader["DefaultLanguage"].ToString() == "TH")
                        {

                            StreamReader streader = new StreamReader(Server.MapPath("~/Template/cfm_th.html"));
                            string readFile = streader.ReadToEnd();
                            //string body = string.Empty;
                            body = readFile;
                            body = body.Replace("%ContractNumber%", reader["ContractNumber"].ToString());
                            body = body.Replace("%OwnerName%", reader["OwnerName"].ToString());
                            body = body.Replace("%MoreAttacthMsg%", "และ เอกสารแสดงคะแนน" );
                            body = body.Replace("%SiteName%", reader["SiteName"].ToString());
                            body = body.Replace("%RoomType%", reader["RoomTypeDesc"].ToString());
                            body = body.Replace("%cInDate%", reader["cInDate"].ToString());
                            body = body.Replace("%cOutDate%", reader["cOutDate"].ToString());
                            body = body.Replace("%PointAmt%", reader["PointAmount"] != DBNull.Value ? Convert.ToDouble(reader["PointAmount"].ToString()).ToString("0") : string.Empty);
                            body = body.Replace("%ChkInTime%", "15.00 น.");
                            switch (Convert.ToInt32(reader["SiteID"]))
                            {
                                case 14:
                                case 15:
                                case 16:
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


                            if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 62 ||
                                   Convert.ToInt32(reader["ReservationSubTypeID"]) == 66)
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
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 146)
                            {
                                if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) > 30)
                                {
                                    CXLPolicyCond += "สำหรับคำขอจองห้องพักที่ได้รับการยืนยัน ถูกร้องขอในช่วงเวลาอันจำกัดและเข้าสู่ช่วงเทศกาล ภายใน 30 วันหรือน้อยกว่าก่อนวันลงทะเบียนเข้าพัก การยกเลิกหรือเปลี่ยนแปลงการจองห้องพักต้องแจ้งให้ทีมบริการสมาชิกทราบเป็นลายลักษณ์อักษรภายใน 24 ชั่วโมงนับจากวันที่ได้รับการยืนยันการเข้าพัก  มิฉะนั้นจะถูกหักคะแนน 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก.";
                                }
                                else if (DateAndTime.DateDiff(DateInterval.Day, Convert.ToDateTime(reader["CConfirmedDate"]), Convert.ToDateTime(reader["InDate"])) <= 30)
                                {
                                    CXLPolicyCond += "สำหรับคำขอจองห้องพักที่ได้รับการยืนยันเข้าสู่ช่วงเทศกาล  การยกเลิกหรือเปลี่ยนแปลงการจองห้องพักต้องแจ้งให้ทีมบริการสมาชิก ทราบเป็นลายลักษณ์อักษรมากกว่า 30 วันล่วงหน้าก่อนวันลงทะเบียนเข้าพัก หากแจ้งน้อยกว่า 30 วัน จะถูกหักคะแนน 50 เปอร์เซ็นต์ และจะถูกหัก 100 เปอร์เซ็นต์ เต็มของคะแนนที่ใช้ในการสำรองห้องพัก หากแจ้งน้อยกว่า 15 วัน ก่อนวันลงทะเบียนเข้าพัก";
                                }
                            }
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 126)
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
                            else if (Convert.ToInt32(reader["ReservationSubTypeID"]) == 81 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 87 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 91 ||
                                      Convert.ToInt32(reader["ReservationSubTypeID"]) == 147)
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


                            subject = reader["CIDMemberNumber"].ToString()+" : จดหมายยืนยันการสำรองห้องพัก - " + reader["SiteName"].ToString();
                        }











                        //var shortURL = Helpers.shortenIt(String.Format(@"http://member.btprivatecollection.com/EmailTracking/Tracking.aspx?email={0}&subject={1}&sentdate={2}&refid={3}&recipent={4}&type={5}&club={6}"
                        //                                              , reader["OwnerEmail"].ToString()
                        //                                              , "ANVC Booking Confirmed - " + reader["SiteName"].ToString()
                        //                                              , DateTime.Now.ToString("dd/MM/yyyy")
                        //                                              , reader["ReservationID"].ToString()
                        //                                              , reader["OwnerName"].ToString()
                        //                                              , "CFMLetter"
                        //                                              , "ANH"
                        //                                              ));
                        //body = body.Replace("%Tracking%", String.Format(@"<img src='{0}'  width=1 height=1>"
                        //                                                , shortURL));

                        //if (i == 0)
                        //    body = body.Replace("%Tracking%", String.Format(@"<img src='http://115.31.143.249/EmailTracking/Tracking.aspx?email={0}&subject={1}&sentdate={2}&refid={3}&recipent={4}&type={5}&club={6}'  width=1 height=1>"
                        //                                                , reader["OwnerEmail"].ToString()
                        //                                                , subject
                        //                                                , DateTime.Now.ToString("dd/MM/yyyy")
                        //                                                , reader["ReservationID"].ToString()
                        //                                                , reader["OwnerName"].ToString()
                        //                                                , "CFMLetter"
                        //                                                , "ANH"
                        //                                                ));
                        //else
                        //   body = body.Replace("%Tracking%", string.Empty);

                        body = body.Replace("%Tracking%", string.Empty);

                        MailMessage Msg = new MailMessage();
                        Msg.Subject = subject;
                        Msg.BodyEncoding = Encoding.GetEncoding("UTF-8");
                        Msg.SubjectEncoding = Encoding.GetEncoding("UTF-8");
                        Msg.From = new MailAddress("anvc@club-memberservices.com", "Angsana Vacation Club");
                        Msg.To.Add(new MailAddress(recipentEmail[i], reader["OwnerName"].ToString())); //reader["OwnerName"].ToString()));
                        //Msg.Bcc.Add(new MailAddress("eakkaphols@lagunaphuket.com", reader["OwnerName"].ToString()));
                        Msg.Bcc.Add(new MailAddress("onlineemailtracking@lagunaphuket.com", reader["OwnerName"].ToString()));
                        //Msg.Bcc.Add(new MailAddress("anvc@angsanavacationclub.com", reader["OwnerName"].ToString()));
                        
                        Msg.Body = body.ToString();
                        Msg.IsBodyHtml = true;


                        //var contentID = "logo";
                        var inlineLogo = new Attachment(Server.MapPath("~/Template/images/ANH_logo.png"));
                        inlineLogo.ContentId = "ANHlogo";
                        inlineLogo.ContentDisposition.Inline = true;
                        inlineLogo.ContentDisposition.DispositionType = System.Net.Mime.DispositionTypeNames.Inline;
                        Msg.Attachments.Add(inlineLogo);


                        ReportDocument rpt = new ReportDocument();

                        /*Add report  ANVC Booking Confirmation Letter*/
                        if (reader["ContractTypeID"].ToString() == "32" || reader["ContractTypeID"].ToString() == "33" || reader["ContractTypeID"].ToString() == "34")
                            rpt.Load(Server.MapPath("~/Report/MemberBookingConfirmation_MRP.rpt"));
                        else if (reader["DefaultLanguage"].ToString() == "EN")
                            rpt.Load(Server.MapPath("~/Report/MemberBookingConfirmation.rpt"));
                        else if (reader["DefaultLanguage"].ToString() == "TH")
                            rpt.Load(Server.MapPath("~/Report/MemberBookingConfirmation (Thai).rpt"));

                        rpt.SetParameterValue("Resort Name", reader["SiteName"].ToString());
                        rpt.SetParameterValue("Confirmation Number", Convert.ToInt32(reader["ReservationNumber"]));
                        rpt.SetParameterValue("Show Room No?","No");
                        rpt.SetDatabaseLogon(Helpers.ReportUserName, Helpers.ReportPassword, Helpers.ReportServer, "TSWDATA");
                        rpt.SetDatabaseLogon(Helpers.ReportUserName, Helpers.ReportPassword, Helpers.ReportServer, Helpers.ReportDatabase);
                        Msg.Attachments.Add(new Attachment(rpt.ExportToStream(ExportFormatType.PortableDocFormat), reader["CIDMemberNumber"].ToString()+"-CFMLetter.pdf"));



                        if (reader["ContractTypeID"].ToString() == "32" || reader["ContractTypeID"].ToString() == "33" || reader["ContractTypeID"].ToString() == "34")
                        {
                            rpt.Load(Server.MapPath("~/Report/PointStatement (Active Contract)_MRP.rpt"));
                            rpt.SetParameterValue("Contract Number", reader["CIDMemberNumber"].ToString());//reader["ContractNumber"].ToString());
                            rpt.SetParameterValue("Show Transaction Summary", false);
                            rpt.SetParameterValue("Show Expire Point", false);
                        }
                        else if (reader["ReservationType"].ToString() == "LHC Member" && reader["ContractNumber"].ToString().Substring(0, 4) == "ANVC")
                        {
                            rpt.Load(Server.MapPath("~/Report/ANVCPointStatement.rpt"));
                            rpt.SetParameterValue("Contract Number", reader["ContractNumber"].ToString());
                            rpt.SetParameterValue("Show Transaction Summary", false);
                        }
                        else if (reader["DefaultLanguage"].ToString() == "EN")
                        {
                            rpt.Load(Server.MapPath("~/Report/PointStatement (Active Contract).rpt"));
                            rpt.SetParameterValue("Contract Number", reader["CIDMemberNumber"].ToString());//reader["ContractNumber"].ToString());
                            rpt.SetParameterValue("Show Transaction Summary", false);
                            rpt.SetParameterValue("Show Expire Point", false);
                        }
                        else if (reader["DefaultLanguage"].ToString() == "TH")
                        {
                            rpt.Load(Server.MapPath("~/Report/PointStatement (Active Contract - Thai).rpt"));
                            rpt.SetParameterValue("Contract Number", reader["CIDMemberNumber"].ToString());//reader["ContractNumber"].ToString());
                            rpt.SetParameterValue("Show Transaction Summary", false);
                            rpt.SetParameterValue("Show Expire Point", false);
                        }

                       
                        rpt.SetDatabaseLogon(Helpers.ReportUserName, Helpers.ReportPassword, Helpers.ReportServer, Helpers.ReportDatabase);
                        rpt.SetDatabaseLogon("btpconline", "admin@local");
               
                        // rpt.Subreports[0].SetDatabaseLogon(Helpers.ReportUserName, Helpers.ReportPassword, Helpers.ReportServer, Helpers.ReportDatabase);

                        Msg.Attachments.Add(new Attachment(rpt.ExportToStream(ExportFormatType.PortableDocFormat), reader["CIDMemberNumber"].ToString()+"-PointStatement.pdf"));



                        /*
                        2	LHC @ Allamanda Laguna Phuket
                        3	LHC @ Angsana Laguna Phuket
                        4	LHC @ Angsana Resort & Spa Bintan
                        6	Laguna Holiday Club Phuket Resort
                        14	LHC @ Twin Peaks Residence
                        15	LHC @ View Talay Residence 6
                        16	LHC @ Boathouse Hua Hin
                        31	LHC Private Pool Villas
                        */
                        switch (Convert.ToInt32(reader["SiteID"]))
                            {   case 2:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["ALMD"]);
                                    break;
                                case 3:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["ANLP"]);
                                    break;
                                case 4:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["ANBT"]);
                                    break;
                                case 6:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["LHCR"]);
                                    break;
                                case 14:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["TWP"]);
                                    break;
                                case 15:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["VTR"]);
                                    break;
                                case 16:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["BHH"]);
                                    break;
                                case 31:
                                    factsheet = Server.MapPath("~/Factsheet/" + ConfigurationManager.AppSettings["PPV"]);
                                    break;
                                default:
                                    break;
                            }

                            Attachment attachment = new Attachment(factsheet);
                            attachment.Name = "Factsheet.pdf";
                            Msg.Attachments.Add(attachment);
                        

                        




                        messageLog += string.Format("|   {0}   |   {1}   |   {2}   |\r\n", reader["OwnerEmail"].ToString()
                                                                                         , Helpers.SendHTMLMail(Msg) == true ? "Pass" : "False"
                                                                                         , DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        messageLog += string.Format("----------------------------------------------------------------------\r\n");

                        Msg = null;






                        //InsertLogCFMLetter(reader);




                    } //end for

                }
                    reader.Dispose();



                    if (File.Exists(@"D:\WebSiteExt\ANHAutomate\LogFile\" + DateTime.Now.ToString("yyyyMMdd") + "_Log_ComfirmationLetter.txt"))
                    {
                        File.Delete(@"D:\WebSiteExt\ANHAutomate\LogFile\" + DateTime.Now.ToString("yyyyMMdd") + "_Log_ComfirmationLetter.txt");
                    }


                    using (StreamWriter log = new StreamWriter(@"D:\WebSiteExt\ANHAutomate\LogFile\" + DateTime.Now.ToString("yyyyMMdd") + "_Log_ComfirmationLetter.txt"))
                    {
                        // Write header of file:
                        log.WriteLine("------------------------------------------------------------");
                        log.WriteLine("|      Script send email automate confirmation letter      |");
                        log.WriteLine("------------------------------------------------------------");
                        log.WriteLine("|  E-mail  |   Status   |   Sent DateTime                  |");
                        log.WriteLine("------------------------------------------------------------");

                        // write transaction to log file
                        log.WriteLine(messageLog);

                        // Close the stream:
                        log.Close();
                    }
                
                
            }

        }

        private static bool InsertLogCFMLetter(SqlDataReader reader )
        {
            string sql = string.Format(@"INSERT INTO [dbo].[ANHOnline_Automate_BookingConfirmationLetter]
           (   [ReservationID]
              ,[OwnerID]
              ,[ContractNumber]
              ,[ReservationNumber]
              ,[CConfirmedDate]
              ,[PrintDate]
              ,[Approval]
              ,[ReservationType]
              ,[ReservationSubType]
              ,[SiteName]
              ,[RoomTypeDesc]
              ,[InDate]
              ,[cIndate]
              ,[cIndate_Chinese]
              ,[OutDate]
              ,[cOutdate]
              ,[cOutdate_Chinese]
              ,[TotalNights]
              ,[cTotalNight]
              ,[cTotalNight_Chinese]
              ,[ReservationStatus]
              ,[rsvn]
              ,[ProspectSalutation]
              ,[ProspectFirstName]
              ,[ProspectLastName]
              ,[ProspectFullName]
              ,[ProspectDefaultEmail]
              ,[DefaultLanguage]
              ,[ProspectDefaultPhone]
              ,[DateCreated]
              ,[ProspectTypeID]
              ,[ProspectType]
              ,[PointAmount]
              ,[ContractType]
              ,[CreatedByUser]
              ,[StaffEmail]
)
     VALUES
           ({0} 
           ,{1}
           ,'{2}'
           ,{3}
           ,'{4}'
           ,'{5}'
           ,'{6}'
           ,'{7}'
           ,'{8}'
           ,'{9}'
           ,'{10}'
           ,'{11}'
           ,'{12}'
           ,'{13}'
           ,'{14}'
           ,'{15}'
           ,'{16}'
           , {17}
           ,'{18}'
           ,'{19}'
           ,'{20}'
           ,'{21}'
           ,'{22}'
           ,'{23}'
           ,'{24}'
           ,'{25}'
           ,'{26}'
           ,'{27}'
           ,'{28}'
           ,'{29}'
           , {30}
           ,'{31}'
           ,'{32}'
           ,'{33}'
           ,'{34}'
           ,'{35}')"
           , reader["ReservationID"].ToString()
           , reader["OwnerID"].ToString()
           , reader["ContractNumber"].ToString()
           , reader["ReservationNumber"].ToString()
           , reader["CConfirmedDate"].ToString()
           , reader["PrintDate"].ToString()
           , reader["Approval"].ToString()
           , reader["ReservationType"].ToString()
           , reader["ReservationSubType"].ToString()
           , reader["SiteName"].ToString()
           , reader["RoomTypeDesc"].ToString()
           , reader["InDate"].ToString()
           , reader["cIndate"].ToString()
           , reader["cIndate_Chinese"].ToString()
           , reader["OutDate"].ToString()
           , reader["cOutdate"].ToString()
           , reader["cOutdate_Chinese"].ToString()
           , reader["TotalNights"].ToString()
           , reader["cTotalNight"].ToString()
           , reader["cTotalNight_Chinese"].ToString()
           , reader["ReservationStatus"].ToString()
           , reader["rsvn"].ToString()
           , reader["ProspectSalutation"].ToString()
           , reader["ProspectFirstName"].ToString()
           , reader["ProspectLastName"].ToString()
           , reader["ProspectFullName"].ToString()
           , reader["ProspectDefaultEmail"].ToString()
           , reader["DefaultLanguage"].ToString()
           , reader["ProspectDefaultPhone"].ToString()
           , reader["DateCreated"].ToString()
           , reader["ProspectTypeID"].ToString()
           , reader["ProspectType"].ToString()
           , reader["PointAmount"].ToString()
           , reader["ContractType"].ToString()
           , reader["CreatedByUser"].ToString()
           , reader["StaffEmail"].ToString()
           );
            using (SqlConnection cn = new SqlConnection(ConnStrInfSRV))
            {
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.CommandType = CommandType.Text;
                cmd.CommandTimeout = 0;
                cn.Open();
                int ret = cmd.ExecuteNonQuery();
                return (ret == 1);
                
            }
        }
        
    }

