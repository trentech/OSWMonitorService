using HtmlAgilityPack;
using Microsoft.CodeAnalysis;
using Serilog;
using System;
using System.Net;
using System.Net.Mail;

namespace OSWMonitorService
{
    public class Worker : BackgroundService
    {
        private Config config;

        public Worker()
        {
            this.config = Config.Get();
        }

        public override Task StartAsync(CancellationToken cancellationToken)
        {
            Log.Information("Starting OSW Monitoring Service.");

            InitDevMode();
            return base.StartAsync(cancellationToken);
        }

        public override Task StopAsync(CancellationToken cancellationToken)
        {
            Log.Information("Stopping OSW Monitoring Service.");

            return base.StopAsync(cancellationToken);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                List<Sensor> list = new List<Sensor>(config.Sensors);
                config.Sensors.Clear();

                foreach (Sensor sensor in list)
                {
                    if (!sensor.Skip)
                    {
                        Log.Information("Getting sensor data on device " + sensor.IP);
                        config.Sensors.Add(GetSensor(sensor.Name, sensor.IP));
                    }
                    else
                    {
                        Log.Information("Skipping device " + sensor.IP);
                    }
                }

                if(config.DataType.Type.Equals(DataType.DataTypes.EXCEL))
                {
                    new Excel(config).AddAll();
                }
                else if (config.DataType.Type.Equals(DataType.DataTypes.ACCESS))
                {
                    new Access(config).AddAll();
                }
                else if (config.DataType.Type.Equals(DataType.DataTypes.MYSQL))
                {
                    new MySQL(config).AddAll();
                }

                int delay = config.DevMode ? 6 : config.Delay;

                await Task.Delay(1000 * 60 * delay, stoppingToken);  
            }
        }

        private Sensor GetSensor(string name, string ip)
        {
            Sensor sensor = new Sensor(name, ip);

            if(config.DevMode)
            {
                if (new Random().Next(0, 100) < 10) //10%
                {
                    Log.Error("[DEVMODE]: Failed to get sensor data on device " + ip);
                    sensor.IsOffline = true;
                }
                else
                {
                    sensor.Temperature = Math.Round(new Random().NextDouble() * (100 - 30) + 30, 2);
                    sensor.Humidity = Math.Round(new Random().NextDouble() * (60 - 30) + 30, 2);
                    sensor.DewPoint = Math.Round(new Random().NextDouble() * (100 - 30) + 30, 2);
                    sensor.IsOffline = false;
                    sensor.IsRecording = true;
                }

                return sensor;
            }

            var url = @"http://" + ip + "/postReadHtml?a=";

            if(!IsOnline(url))
            {
                Log.Error("Failed to get sensor data on device " + ip);

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

                smtpClient.Send(message);

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

        private void InitDevMode()
        {
            if(config.DevMode)
            {
                for(int i = 1; i < 10; i++)
                {
                    Sensor sensor = new Sensor("Test" + i, "10.0.0." + i);
                    sensor.Skip = false;
                    config.Sensors.Add(sensor);
                }
            }
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