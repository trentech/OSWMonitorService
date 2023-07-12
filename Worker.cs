using HtmlAgilityPack;
using Serilog;
using System.Diagnostics;
using System.Net;
using OSWMonitorService.DataTypes;
using OSWMonitorService.JSON;
using OSWMonitorService.Properties;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.Formula.PTG;

namespace OSWMonitorService
{
    public class Worker : BackgroundService
    {
        public override Task StartAsync(CancellationToken cancellationToken)
        {
            Log.Information("Starting OSW Monitoring Service.");

            Config config = Config.Get();

            if (config.DataType.Type.Equals(DataType.DataTypes.ACCESS))
            {
                string path = Path.Combine(config.DataType.Path, config.DataType.Name) + ".accdb";

                if (!File.Exists(path))
                {
                    File.WriteAllBytes(path, Resources.database);
                }
            }

            return base.StartAsync(cancellationToken);
        }

        public override Task StopAsync(CancellationToken cancellationToken)
        {
            Log.Information("Stopping OSW Monitoring Service.");

            return base.StopAsync(cancellationToken);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            Config config = Config.Get();
            
            while (!stoppingToken.IsCancellationRequested)
            {
                int delay = config.Delay > 60 ? 60 : config.Delay;
                int minute = 0;
                int current = DateTime.Now.Minute;
                bool run = false;

                for (int i = 0; i < (60 / delay); i++)
                {
                    minute = minute + delay;
                    minute = minute == 60 ? 0 : minute;

                    if (minute == current) run = true;
                }

                if (run)
                {
                    Log.Information("Executing");

                    config = Config.Get();

                    foreach (Sensor sensor in Sensor.All())
                    {
                        if (sensor.Skip)
                        {
                            Log.Information("[" + sensor.IP + "] Skipped");
                            continue;
                        }

                        _ = NewTask(config, sensor);
                    }
                }

                DateTime now = DateTime.Now;
                await Task.Delay((((now.Second > 30 ? 120 : 60) - now.Second) * 1000 - now.Millisecond) + 500, stoppingToken);
            }
        }

        private Task NewTask(Config config, Sensor sensor)
        {
            return Task.Run(() =>
            {
                Log.Information("[" + sensor.IP + "] Getting sensor data");

                bool wasOffline = false;

                if (!sensor.IsOnline)
                {
                    wasOffline = true;
                }

                sensor = ParseSensor(sensor);

                Stopwatch stopWatch = Stopwatch.StartNew();

                while (!sensor.IsOnline)
                {
                    if (stopWatch.Elapsed.TotalSeconds >= 59)
                    {
                        Log.Error("[" + sensor.IP + "] Sensor offline. Timing out");

                        if (!wasOffline)
                        {
                            Utils.SendEmail(config.Email, sensor.Recipients, "[" + sensor.Name + "] Sensor Alarm - Connection Timeout", "Name: " + sensor.Name + Environment.NewLine + "IP: " + sensor.IP + Environment.NewLine + "Online: " + sensor.IsOnline);
                        }
                        
                        stopWatch.Stop();
                        sensor.Save();
                        return;
                    }
                    else
                    {
                        sensor = ParseSensor(sensor);
                    }
                }

                stopWatch.Stop();

                if (wasOffline && sensor.IsOnline)
                {
                    Log.Information("[" + sensor.IP + "] Sensor online");

                    Utils.SendEmail(config.Email, sensor.Recipients, "[" + sensor.Name + "] Sensor Alarm - Online", "Name: " + sensor.Name + Environment.NewLine + "IP: " + sensor.IP + Environment.NewLine + "Online: " + sensor.IsOnline);
                }

                if (sensor.IsOnline)
                {
                    if (sensor.HumidityWarning != 0 && sensor.Humidity > sensor.HumidityWarning && sensor.Humidity < sensor.HumidityLimit)
                    {
                        Log.Warning("[" + sensor.IP + "] Humidity Threshold Warning!");

                        Utils.SendEmail(config.Email, sensor.Recipients, "[" + sensor.Name + "] Sensor Alarm - Humidity Threshold Warning", "Name: " + sensor.Name + Environment.NewLine + "IP: " + sensor.IP + Environment.NewLine + "Temperature: " + sensor.Temperature
                            + Environment.NewLine + "Humidity: " + sensor.Humidity);
                    }

                    if (sensor.HumidityLimit != 0 && sensor.Humidity > sensor.HumidityLimit)
                    {
                        Log.Warning("[" + sensor.IP + "] Humidity Threshold Reached!");

                        Utils.SendEmail(config.Email, sensor.Recipients, "[" + sensor.Name + "] Sensor Alarm - Humidity Threshold Reached", "Name: " + sensor.Name + Environment.NewLine + "IP: " + sensor.IP + Environment.NewLine + "Temperature: " + sensor.Temperature
                            + Environment.NewLine + "Humidity: " + sensor.Humidity);
                    }

                    if (sensor.TemperatureLimit != 0 && sensor.Temperature > sensor.TemperatureLimit)
                    {
                        Log.Warning("[" + sensor.IP + "] Temperature Threshold Reached!");

                        Utils.SendEmail(config.Email, sensor.Recipients, "[" + sensor.Name + "] Sensor Alarm - Temperature Threshold Reached", "Name: " + sensor.Name + Environment.NewLine + "IP: " + sensor.IP + Environment.NewLine + "Temperature: " + sensor.Temperature
                            + Environment.NewLine + "Humidity: " + sensor.Humidity);
                    }

                    // 5 DEGREES APART
                    if ((sensor.Temperature > sensor.DewPoint ? sensor.Temperature - sensor.DewPoint : sensor.DewPoint - sensor.Temperature) <= 5)
                    {
                        Log.Warning("[" + sensor.IP + "] Condensation Warning!");

                        // THE DIDN"T ASK FOR THIS, BUT HAVE A FEELING THEY WILL. DISABLED FOR NOW.
                        //    Utils.SendEmail(config.Email, s.Recipients, "[" + s.Name + "] Sensor Alarm - Condensation Warning", "Name: " + s.Name + Environment.NewLine + "IP: " + s.IP + Environment.NewLine + "Temperature: " + s.Temperature
                        //        + Environment.NewLine + "Humidity: " + s.Humidity + Environment.NewLine + "Dew Point: " + s.DewPoint);
                    }
                }

                if (config.DataType.Type.Equals(DataType.DataTypes.ACCESS))
                {
                    new Access(config).AddEntry(sensor);
                }
                else if (config.DataType.Type.Equals(DataType.DataTypes.MYSQL))
                {
                    new MySQL(config).AddEntry(sensor);
                }
                else
                {
                    Log.Error("Invalid Database Type in config");
                }

                sensor.Save();
            });
        }

        private Sensor ParseSensor(Sensor sensor)
        {
            var url = @"http://" + sensor.IP + "/postReadHtml?a=";

            HtmlWeb web = new HtmlWeb();

            web.PreRequest = delegate (HttpWebRequest webRequest)
            {
                webRequest.Timeout = 10000;
                return true;
            };

            HtmlDocument htmlDoc = null;
            try
            {
                htmlDoc = web.Load(url);
                sensor.IsOnline = true;
            }
            catch(WebException ex) 
            {
                Log.Error("[" + sensor.IP + "] " + ex.Message);
                sensor.IsOnline = false;

                return sensor;
            }

            string data = htmlDoc.DocumentNode.OuterHtml;

            string[] lines = data.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            sensor.DateTime = DateTime.Now;

            string[] temp = lines[1].Substring(2).Split(" ");
            sensor.Temperature = Double.Parse(temp[temp.Length - 2]);

            string[] humid = lines[2].Substring(2).Split(" ");
            sensor.Humidity = Double.Parse(humid[humid.Length - 2]);

            string[] dew = lines[3].Substring(2).Split(" ");
            sensor.DewPoint = Double.Parse(dew[dew.Length - 2]);

            return sensor;
        }
    }
}