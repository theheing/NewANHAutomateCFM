using System;
using System.Collections.Generic;
using System.Web;
using System.Globalization;

using System.IO;
using System.Net.Mail;
using System.Web.UI;
using System.Linq;
using System.Configuration;
using System.Web.Configuration;
using System.Net.Configuration;



using Google;
using Google.Apis.Services;
using Google.Apis.Urlshortener.v1;
using Google.Apis.Urlshortener.v1.Data;


namespace Automate.Utilities
{
    public static partial class Helpers
    {
        private static IFormatProvider enCulture = new CultureInfo("en-US");
        private static IFormatProvider thCulture = new CultureInfo("th-TH");
        //private string reportServer = ConfigurationManager.AppSettings["ReportServer"];
        //private string reportDatabase = ConfigurationManager.AppSettings["ReportDatabase"];
        //private string reportUserName = ConfigurationManager.AppSettings["ReportUserName"];
        //private string reportPassword = ConfigurationManager.AppSettings["ReportPassword"];


        public static IFormatProvider ENCulture
        {
            get { return enCulture; }
        }
        public static IFormatProvider THCulture
        {
            get { return thCulture; }
        }
        public static string ReportServer
        {
            get { return ConfigurationManager.AppSettings["ReportServer"]; }
        }
        public static string ReportDatabase
        {
            get { return ConfigurationManager.AppSettings["ReportDatabase"]; }
        }
        public static string ReportUserName
        {
            get { return ConfigurationManager.AppSettings["ReportUserName"]; }
        }
        public static string ReportPassword
        {
            get { return ConfigurationManager.AppSettings["ReportPassword"]; }
        }

        public static bool SendHTMLMail(MailMessage msg)
        {
            //Configuration configurationFile = WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Server.MapPath("~/web.config"));
            //Configuration configurationFile = WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Server.MapPath("~/web.config"));
            //MailSettingsSectionGroup mailSettings = configurationFile.GetSectionGroup("system.net/mailSettings") as MailSettingsSectionGroup;

            //if (mailSettings != null)
            //{
            //    try
            //    {
            //        SmtpClient smtp = new SmtpClient();
            //        smtp.Host = mailSettings.Smtp.Network.Host;
            //        smtp.Port = mailSettings.Smtp.Network.Port;
            //        smtp.Credentials = new System.Net.NetworkCredential(mailSettings.Smtp.Network.UserName
            //                                                           , mailSettings.Smtp.Network.Password);
            //        smtp.EnableSsl = false;
            //        smtp.Send(msg);
            //    }
            //    catch (Exception ex)
            //    {
            //        throw new Exception("Error Infomation ", ex);
            //    }
            //}

            Configuration configurationFile = WebConfigurationManager.OpenWebConfiguration("~/web.config");
            MailSettingsSectionGroup mailSettings = configurationFile.GetSectionGroup("system.net/mailSettings") as MailSettingsSectionGroup;

            bool success = false;
            System.Net.ServicePointManager.ServerCertificateValidationCallback =
                delegate(object s, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
                { return true; };
            SmtpClient smtp = new SmtpClient();
            smtp.Host = mailSettings.Smtp.Network.Host;
            smtp.UseDefaultCredentials = mailSettings.Smtp.Network.DefaultCredentials;
            smtp.EnableSsl = mailSettings.Smtp.Network.EnableSsl;
            smtp.Credentials = new System.Net.NetworkCredential(mailSettings.Smtp.Network.UserName, mailSettings.Smtp.Network.Password);
            smtp.Port = mailSettings.Smtp.Network.Port;

            try
            {
                smtp.Send(msg);
                success = true;
            }
            catch
            {
                success = false;
            }
            return success;

        }

        public static string shortenIt(string url)
        {
            UrlshortenerService service = new UrlshortenerService(new BaseClientService.Initializer()
            {
                ApiKey = "AIzaSyALE8yFSn-PIjb3IAEsn1y43OCqtHynH-k",
                ApplicationName = "EmailTracking",
            });

            var m = new Google.Apis.Urlshortener.v1.Data.Url();
            m.LongUrl = url;
            return service.Url.Insert(m).Execute().Id;
        }

        public static string unShortenIt(string url)
        {
            UrlshortenerService service = new UrlshortenerService(new BaseClientService.Initializer()
            {
                ApiKey = "AIzaSyALE8yFSn-PIjb3IAEsn1y43OCqtHynH-k",
                ApplicationName = "Daimto URL shortener Sample",
            });
            return service.Url.Get(url).Execute().LongUrl;
        }

        public static bool RemoteCertificateValidationCB(Object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,

System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {

            return true;

        }



    }
}
