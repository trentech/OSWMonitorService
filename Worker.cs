using HtmlAgilityPack;
using Microsoft.CodeAnalysis;
using System.Net;
using System.Net.Mail;

namespace OSWMontiorService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> logger;

        private Config config;

        public Worker(ILogger<Worker> logger)
        {
            this.logger = logger;
            this.config = Config.Get();
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            logger.LogInformation("Starting OSW Monitoring Service: {time}", DateTime.Now);

            while (!stoppingToken.IsCancellationRequested)
            {
                List<Sensor> list = new List<Sensor>(config.Sensors);
                config.Sensors.Clear();

                foreach (Sensor sensor in list)
                {
                    if (!sensor.Skip)
                    {
                        logger.LogInformation("Getting sensor data on device " + sensor.IP + ": {time}", DateTime.Now);
                        config.Sensors.Add(GetSensor(sensor.Name, sensor.IP));
                    }
                    else
                    {
                        logger.LogInformation("Skipping device " + sensor.IP + ": {time}", DateTime.Now);
                    }
                }

                if (config.DataType == "Excel")
                {
                    new Excel(logger, config).AddAll();
                }
                else if (config.DataType == "Access")
                {
                    new Access(logger, config).AddAll();
                }
                else if (config.DataType == "MySQL")
                {
                    // NEED CONFIG OPTIONS FOR NULL VALUES
                    new MySQL(logger, config, null, null, null, null).AddAll();
                }

                await Task.Delay(1000 * 60 * config.Delay, stoppingToken);
            }
        }

        private Sensor GetSensor(string name, string ip)
        {
            Sensor sensor = new Sensor(name, ip);
            sensor.Skip = false;

            var url = @"http://" + ip + "/postReadHtml?a=";

            if(!IsOnline(url))
            {
                logger.LogError("Failed to get sensor data on device " + ip + ": {time}", DateTime.Now);

                Mail mail = new Mail();

                SmtpClient smtpClient = new SmtpClient(mail.STMP) 
                {
                    Port = mail.Port,
                    EnableSsl = mail.SSL,
                };

                MailMessage message = new MailMessage();
                message.From = new MailAddress(mail.From);

                foreach(string address in mail.Recipients)
                {
                    message.To.Add(new MailAddress(address));
                }

                message.Subject = "OSW Sensor Offline";
                message.Body = "[" + sensor.Name + " : " + sensor.IP + "] Sensor Offline";

            //    smtpClient.Send(message);

                sensor.IsOffline = true;

                return sensor;
            }

            HtmlWeb web = new HtmlWeb();

            HtmlDocument htmlDoc = web.Load(url);

            string data = htmlDoc.DocumentNode.OuterHtml;

            string[] lines = data.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            string[] date = lines[0].Substring(2).Split(" ");
            sensor.Time = TimeOnly.Parse(date[3].PadLeft(2, '0') + ":" + date[4].PadLeft(2, '0') + ":" + date[5].PadLeft(2, '0'));
            sensor.Date = DateOnly.Parse((Int32.Parse(date[1]) - 1) + "-" + date[2] + "-" + date[0]);

            string[] temp = lines[1].Substring(2).Split(" ");
            sensor.Temperature = Double.Parse(temp[temp.Length - 2]);

            string[] humid = lines[2].Substring(2).Split(" ");
            sensor.Humidity = Double.Parse(humid[humid.Length - 2]);

            string[] dew = lines[3].Substring(2).Split(" ");
            sensor.DewPoint = Double.Parse(dew[dew.Length - 2]);

            string recordUnformatted = lines[4].Substring(2).Split(" ")[1];
            sensor.IsRecording = recordUnformatted.Equals("ON") ? true : false;

            return sensor;
        }

        private bool IsOnline(string url)
        {
            try
            {
#pragma warning disable SYSLIB0014 // Type or member is obsolete
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
#pragma warning restore SYSLIB0014 // Type or member is obsolete
                request.Timeout = 3000;
                request.AllowAutoRedirect = false;
                request.Method = "GET";

                using (var response = request.GetResponse())
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}