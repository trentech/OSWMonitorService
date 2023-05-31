using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Serilog;

namespace OSWMonitorService
{
    internal class Utils
    {
        public static void SendEmail(Mail mail, string subject, string body)
        {
            SmtpClient smtpClient = new SmtpClient(mail.STMP)
            {
                Port = mail.Port,
                EnableSsl = mail.SSL,
                UseDefaultCredentials = true
            };

            MailMessage message = new MailMessage();
            message.From = new MailAddress(mail.From);

            foreach (string address in mail.Recipients)
            {
                MailAddress email = new MailAddress(address);
                Log.Information("Sending email to " + email.Address);
                message.To.Add(email);
            }

            message.Subject = subject;
            message.Body = body;

            try
            {
                smtpClient.Send(message);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Email notification failed");
            }
        }
    }
}
