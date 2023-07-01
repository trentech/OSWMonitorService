using System.Net.Mail;
using Serilog;
using OSWMonitorService.JSON;

namespace OSWMonitorService
{
    internal class Utils
    {
        public static void SendEmail(Mail mail, List<string> recipients, string subject, string body, bool dev)
        {
            if (recipients.Count == 0)
            {
                return;
            }

            SmtpClient smtpClient = new SmtpClient(mail.STMP)
            {
                Port = mail.Port,
                EnableSsl = mail.SSL,
                UseDefaultCredentials = true
            };

            MailMessage message = new MailMessage();
            message.From = new MailAddress(mail.From);

            foreach (string address in recipients)
            {
                MailAddress email = new MailAddress(address);
                Log.Information("Sending email to " + email.Address + " Subject: " + subject);
                message.To.Add(email);
            }

            message.Subject = subject;
            message.Body = body;

            if (!dev)
            {
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
}
