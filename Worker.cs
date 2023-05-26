using HtmlAgilityPack;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Diagnostics;

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
                        GetSensor(sensor.Name, sensor.IP);
                    }
                    else
                    {
                        logger.LogInformation("Skipping device " + sensor.IP + ": {time}", DateTime.Now);
                    }
                }

                WriteData();

                await Task.Delay(1000 * 60 * config.Delay, stoppingToken);
            }
        }

        private void GetSensor(string name, string ip)
        {
            Sensor sensor = new Sensor(name, ip);
            sensor.Skip = false;

            var url = @"http://" + ip + "/postReadHtml?a=";

            if(!IsOnline(url))
            {
                logger.LogError("Failed to get sensor data on device " + ip + ": {time}", DateTime.Now);

                // ADD LOGIC HERE TO SEND OUT NOTIFICATIONS
                // ALL THIS NEEDS UPDATING

              //SmtpClient smtpClient = new SmtpClient("arvosgroup-com01c.mail.protection.outlook.com")
                SmtpClient smtpClient = new SmtpClient("10.164.123.206") 
                {
                    Port = 25,
                    EnableSsl = false,
                };

                smtpClient.Send("IT.US@ljungstrom.com", "terrence.monroe@ljungstrom.com", "OSW Sensor Offline", "[" + sensor.Name + " : " + sensor.IP + "] Sensor Offline");

                sensor.IsOffline = true;
                config.Sensors.Add(sensor);
                return;
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

            config.Sensors.Add(sensor);
        }

        public void CreateSpreadSheet(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                XSSFWorkbook workbook = new XSSFWorkbook();

                foreach (Sensor sensor in config.Sensors)
                {
                    ISheet sheet = workbook.CreateSheet(sensor.IP);
                    sheet.DefaultColumnWidth = 15;

                    IFont font = workbook.CreateFont();
                    font.IsBold = true;
                    font.FontHeightInPoints = 16;

                    ICellStyle style = workbook.CreateCellStyle();
                    style.Alignment = HorizontalAlignment.Center;
                    style.VerticalAlignment = VerticalAlignment.Center;
                    style.SetFont(font);

                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 5));
                    IRow row = sheet.CreateRow(0);
                    row.Height = 20 * 20;
                    ICell cell = row.CreateCell(0);
                    cell.SetCellValue(sensor.Name);
                    cell.CellStyle = style;

                    font = workbook.CreateFont();
                    font.IsBold = true;
                    font.Color = HSSFColor.White.Index;

                    style = workbook.CreateCellStyle();
                    style.FillForegroundColor = HSSFColor.Grey50Percent.Index;
                    style.FillPattern = FillPattern.SolidForeground;
                    style.Alignment = HorizontalAlignment.Center;
                    style.SetFont(font);

                    row = sheet.CreateRow(1);

                    cell = row.CreateCell(0);
                    cell.CellStyle = style;
                    cell.SetCellValue("Temperature");

                    cell = row.CreateCell(1);
                    cell.CellStyle = style;
                    cell.SetCellValue("Humidity");

                    cell = row.CreateCell(2);
                    cell.CellStyle = style;
                    cell.SetCellValue("Dew");

                    cell = row.CreateCell(3);
                    cell.CellStyle = style;
                    cell.SetCellValue("Recording");

                    cell = row.CreateCell(4);
                    cell.CellStyle = style;
                    cell.SetCellValue("Date");

                    cell = row.CreateCell(5);
                    cell.CellStyle = style;
                    cell.SetCellValue("Time");
                }

                workbook.Write(stream);
                stream.Close();
            }
        }

        public void WriteData()
        {
          //string tempFile =  Path.Combine(Config.PATH, DateTime.Now.Month + "-" + DateTime.Now.Year + ".xlsx");
            string tempFile = Path.Combine(Config.PATH, "temp.xlsx");

            if(!File.Exists(Path.Combine(config.Destination, "OSW Sensors.xlsx")))
            {
                if (File.Exists(tempFile))
                {
                    File.Move(tempFile, Path.Combine(Config.PATH, "temp" + " " + DateOnly.FromDateTime(DateTime.Now).ToString() + ".xlsx"));
                }
            }

            if (!File.Exists(tempFile))
            {
                CreateSpreadSheet(tempFile);
            }

            XSSFWorkbook workbook;

            while(IsFileLocked(tempFile))
            {
                logger.LogWarning("["+ tempFile + "]: File is locked and cannot be written to. Trying again in 30 Seconds: {time}", DateTime.Now);
                Thread.Sleep(30000);
            }

            using (FileStream stream = new FileStream(tempFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(stream);
                stream.Close();
            }

            foreach (Sensor sensor in config.Sensors)
            {
                if (sensor.Skip)
                {
                    continue;
                }

                // NEED LOGIC FOR DETECTING AND CREATING NEW SHEETS IF SENSORS ARE ADDED

                ISheet sheet = workbook.GetSheet(sensor.IP);

                IRow row = sheet.CreateRow(sheet.LastRowNum + 1);

                if(sensor.IsOffline)
                {
                    ICellStyle style = workbook.CreateCellStyle();
                    style.FillForegroundColor = HSSFColor.Red.Index;
                    style.FillPattern = FillPattern.SolidForeground;

                    ICell cell = row.CreateCell(0);
                    cell.CellStyle = style;
                    cell.SetCellValue("NA");

                    cell = row.CreateCell(1);
                    cell.CellStyle = style;
                    cell.SetCellValue("NA");

                    cell = row.CreateCell(2);
                    cell.CellStyle = style;
                    cell.SetCellValue("NA");

                    cell = row.CreateCell(3);
                    cell.CellStyle = style;
                    cell.SetCellValue("NA");

                    row.CreateCell(4).SetCellValue(DateOnly.FromDateTime(DateTime.Now).ToString());
                    row.CreateCell(5).SetCellValue(TimeOnly.FromDateTime(DateTime.Now).ToString());
                } else
                {
                    row.CreateCell(0).SetCellValue(sensor.Temperature);
                    row.CreateCell(1).SetCellValue(sensor.Humidity);
                    row.CreateCell(2).SetCellValue(sensor.DewPoint);
                    row.CreateCell(3).SetCellValue(sensor.IsRecording);
                    row.CreateCell(4).SetCellValue(sensor.Date.ToString());
                    row.CreateCell(5).SetCellValue(sensor.Time.ToString());
                }
            }

            using (FileStream stream = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
                stream.Close();
            }

            CopyDataToPath(tempFile, config.Destination);
        }

        private async void CopyDataToPath(string source, string destination)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            TimeOnly time = TimeOnly.FromDateTime(DateTime.Now);

          //string file = Path.Combine(destination, "OSW Sensors " + DateTime.Now.Month + "-" + DateTime.Now.Year + ".xlsx");
            string file = Path.Combine(destination, "OSW Sensors.xlsx");

            await Task.Run(() => {
                if (File.Exists(file))
                {
                    while (IsFileLocked(file))
                    {
                        logger.LogWarning("[" + file + "]: File is locked and cannot be written to. Trying again in 30 Seconds: {time}", DateTime.Now);
                        Thread.Sleep(30000);

                        TimeSpan timeSpan = stopWatch.Elapsed;

                        if(timeSpan.TotalHours >= 1)
                        {

                          // NEEDS UPDATING

                          //SmtpClient smtpClient = new SmtpClient("arvosgroup-com01c.mail.protection.outlook.com")
                            SmtpClient smtpClient = new SmtpClient("10.164.123.206")            
                            {
                                Port = 25,
                                EnableSsl = false,
                            };

                            smtpClient.Send("IT.US@ljungstrom.com", "terrence.monroe@ljungstrom.com", "OSW Sensor File Locked", "[" + file + "] File has been locked from editing for over an hour.");

                            logger.LogError("[" + file + "]: File is locked and cannot be written to: {time}", DateTime.Now);
                            return;
                        }
                    }

                    File.Delete(file);
                }

                logger.LogInformation("[" + file + "]: Writing to file: {time}", DateTime.Now);
                File.Copy(source, file);
            });
        }

        protected virtual bool IsFileLocked(string path)
        {
            try
            {
                using (FileStream stream = new FileInfo(path).Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                return true;
            }

            return false;
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