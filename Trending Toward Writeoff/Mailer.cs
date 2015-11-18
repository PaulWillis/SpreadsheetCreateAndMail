using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;
using System.Collections.Specialized;
using System.Configuration;

namespace Trending_Toward_Writeoff
{
    public class Mailer
    {

        private static NameValueCollection appConfig = ConfigurationManager.AppSettings;

        public Mailer(string ToCSV, string Subject, string Body, string FileToAttach_IncludingPath)
        {
            try
            {

                string from = appConfig["MailFromAddress"];
                string to = ToCSV;
                string bcc = "";
                string serveraddress =appConfig["MailServerIp"];
                 
                MailMessage mail = new MailMessage();
                  
                mail.From = new MailAddress(from);
                mail.To.Add(to);
                 

                //set the content
                mail.Subject = Subject;
                mail.IsBodyHtml = true;
                mail.Body = Body;
                 
                Attachment data = new Attachment(FileToAttach_IncludingPath, MediaTypeNames.Application.Octet);

                // Add time stamp information for the file.
                ContentDisposition disposition = data.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(FileToAttach_IncludingPath);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(FileToAttach_IncludingPath);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(FileToAttach_IncludingPath);

                mail.Attachments.Add(data); 

                //send
                SmtpClient smtp = new SmtpClient(serveraddress);
                smtp.Send(mail);
            }
            catch (Exception ex)
            { 
                throw ex;
            }
        }
    }
}
