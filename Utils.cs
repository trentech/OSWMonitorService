using System.Net.Mail;
using Serilog;
using OSWMonitorService.JSON;
using System.Net;
using Newtonsoft.Json;

namespace OSWMonitorService
{
    internal class Utils
    {

        public static List<Sensor> GetSensors()
        {
            List<Sensor> list = new List<Sensor>();

            string path = Path.Combine(Config.PATH, "Sensors");

            foreach (var sensor in Directory.GetFiles(path, "*.json"))
            {
                try
                {
                    Sensor s = JsonConvert.DeserializeObject<Sensor>(File.ReadAllText(sensor));
                    if (s != null) { list.Add(s); } else { Log.Error("Cannot deserialize sensor from JSON. It is null"); } 
                } catch (Exception e)
                {
                    Log.Error("Cannot deserialize sensor from JSON", e);
                }
            }

            return list;
        }

        public static Sensor GetSensor(string IP)
        {
            string path = Path.Combine(Config.PATH, "Sensors");
            string sensor = Path.Combine(path, IP + ".json");

            if (File.Exists(sensor))
            {
                try
                {
                    return JsonConvert.DeserializeObject<Sensor>(File.ReadAllText(sensor));
                }
                catch (Exception e)
                {
                    Log.Error("Cannot deserialize sensor from JSON", e);
                }
            }

            return null;
        }

        public static void SendEmail(Mail mail, List<string> recipients, string subject, string body)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

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

            if (mail.Username != "" && mail.Password != "")
            {
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential(mail.Username, mail.Password);
            }

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
