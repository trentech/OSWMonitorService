using HtmlAgilityPack;
using Serilog;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net;
using OSWMonitorService.DataTypes;
using OSWMonitorService.JSON;
using OSWMonitorService.Properties;

namespace OSWMonitorService
{
    public class Worker : BackgroundService
    {
        private List<Sensor> devSensors = new List<Sensor>();

        public override Task StartAsync(CancellationToken cancellationToken)
        {
            Log.Information("Starting OSW Monitoring Service.");

            Config config = Config.Get();

            foreach(Sensor sensor in Utils.GetSensors())
            {
                if (new Random().Next(100) < 5)
                {
                    sensor.IsOnline = false;
                }

                devSensors.Add(sensor);
            }

            if (config.DataType.Type.Equals(DataType.DataTypes.ACCESS))
            {
                string dbFile = Path.Combine(config.DataType.Path, config.DataType.Name) + ".accdb";

                if (!File.Exists(dbFile))
                {
                    File.WriteAllBytes(dbFile, Resources.database);
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
                    Execute(config);
                    config = Config.Get();
                }

                DateTime now = DateTime.Now;
                await Task.Delay((((now.Second > 30 ? 120 : 60) - now.Second) * 1000 - now.Millisecond) + 500, stoppingToken);
            }
        }

        private void Execute(Config config)
        {
            List<Sensor> list = config.DevMode ? devSensors : Utils.GetSensors();

            foreach (Sensor sensor in list)
            {
                if (sensor.Skip)
                {
                    Log.Information("[" + sensor.IP + "] Skipping device");
                    continue;
                }

                Task.Run(() => 
                {
                    Log.Information("[" + sensor.IP + "] Getting sensor data");

                    bool wasOffline = false;

                    if(!sensor.IsOnline)
                    {
                        wasOffline = true;
                    }

                    Sensor s = ParseSensor(sensor, config.DevMode);

                    Stopwatch stopWatch = Stopwatch.StartNew();

                    while (!s.IsOnline)
                    {
                        Log.Warning("[" + s.IP + "] Sensor offline. Trying again...");

                        if (stopWatch.Elapsed.TotalSeconds >= 59000)
                        {
                            Log.Error("[" + s.IP + "] Sensor offline. Timing out");

                            Utils.SendEmail(config.Email, s.Recipients, "[" + s.Name + "] Sensor Alarm - Connection Timeout", "Name: " + s.Name + Environment.NewLine + "IP: " + s.IP + Environment.NewLine + "Online: " + s.IsOnline);

                            stopWatch.Stop();
                            return;
                        }
                        else
                        {
                            s = ParseSensor(s, config.DevMode);
                        }
                    }

                    stopWatch.Stop();

                    if (wasOffline)
                    {
                        Log.Information("[" + s.IP + "] Sensor online");

                        Utils.SendEmail(config.Email, s.Recipients, "[" + s.Name + "] Sensor Alarm - Online", "Name: " + s.Name + Environment.NewLine + "IP: " + s.IP + Environment.NewLine + "Online: " + s.IsOnline);
                    }

                    if (config.DataType.Type.Equals(DataType.DataTypes.ACCESS))
                    {
                        new Access(config).AddEntry(s);
                    }
                    else if (config.DataType.Type.Equals(DataType.DataTypes.MYSQL))
                    {
                        new MySQL(config).AddEntry(s);
                    }
                    else
                    {
                        Log.Error("Invalid Database Type in config");
                    }

                    if (s.HumidityWarning != 0 && s.Humidity > s.HumidityWarning && s.Humidity < s.HumidityLimit)
                    {
                        Log.Warning("[" + s.IP + "] Humidity Threshold Warning!");

                        Utils.SendEmail(config.Email, s.Recipients, "[" + s.Name + "] Sensor Alarm - Humidity Threshold Warning", "Name: " + s.Name + Environment.NewLine + "IP: " + s.IP + Environment.NewLine + "Temperature: " + s.Temperature
                            + Environment.NewLine + "Humidity: " + s.Humidity);
                    }

                    if (s.HumidityLimit != 0 && s.Humidity > s.HumidityLimit)
                    {
                        Log.Warning("[" + s.IP + "] Humidity Threshold Reached!");

                        Utils.SendEmail(config.Email, s.Recipients, "[" + s.Name + "] Sensor Alarm - Humidity Threshold Reached", "Name: " + s.Name + Environment.NewLine + "IP: " + s.IP + Environment.NewLine + "Temperature: " + s.Temperature 
                            + Environment.NewLine + "Humidity: " + s.Humidity);
                    }

                    if (s.Temperature > s.TemperatureLimit)
                    {
                        Log.Warning("[" + s.IP + "] Temperature Threshold Reached!");

                        Utils.SendEmail(config.Email, s.Recipients, "[" + s.Name + "] Sensor Alarm - Temperature Threshold Reached", "Name: " + s.Name + Environment.NewLine + "IP: " + s.IP + Environment.NewLine + "Temperature: " + s.Temperature
                            + Environment.NewLine + "Humidity: " + s.Humidity);
                    }

                    // 5 DEGREES APART
                    if((s.Temperature > s.DewPoint ? s.Temperature - s.DewPoint : s.DewPoint - s.Temperature) <= 5)
                    {
                        Log.Warning("[" + s.IP + "] Condensation Warning!");

                    // THE DIDN"T ASK FOR THIS, BUT HAVE A FEELING THEY WILL. DISABLED FOR NOW.
                    //    Utils.SendEmail(config.Email, s.Recipients, "[" + s.Name + "] Sensor Alarm - Condensation Warning", "Name: " + s.Name + Environment.NewLine + "IP: " + s.IP + Environment.NewLine + "Temperature: " + s.Temperature
                    //        + Environment.NewLine + "Humidity: " + s.Humidity + Environment.NewLine + "Dew Point: " + s.DewPoint);
                    }

                    s.Save();
                });
            }
        }

        private Sensor ParseSensor(Sensor sensor, bool dev)
        {
            if (dev)
            {
                if (!sensor.IsOnline)
                {
                    if (new Random().Next(100) < 5)
                    {
                        sensor.IsOnline = true;
                    }
                }

                sensor.Temperature = Math.Round(new Random().NextDouble() * (105 - 30) + 30, 2);
                sensor.Humidity = Math.Round(new Random().NextDouble() * (60 - 10) + 10, 2);
                sensor.DewPoint = Math.Round(new Random().NextDouble() * (105 - 30) + 30, 2);

                return sensor;
            }

            var url = @"http://" + sensor.IP + "/postReadHtml?a=";

            HtmlWeb web = new HtmlWeb();

            web.PreRequest = delegate (HttpWebRequest webRequest)
            {
                webRequest.Timeout = 10000;
                return true;
            }; 

            HtmlDocument htmlDoc = web.Load(url);

            if(htmlDoc == null)
            {
                Log.Error("[" + sensor.IP + "] Unable to scrape data");
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