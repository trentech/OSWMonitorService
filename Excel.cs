using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.Net.Mail;

namespace OSWMontiorService
{
    public class Excel
    {

        public static void AddAll(ILogger<Worker> logger)
        {
            Config config = Config.Get();

            string tempFile = Path.Combine(Config.PATH, "temp.xlsx");

            if (!File.Exists(Path.Combine(config.Destination, "OSW Sensors.xlsx")))
            {
                if (File.Exists(tempFile))
                {
                    File.Move(tempFile, Path.Combine(Config.PATH, "temp" + " " + DateOnly.FromDateTime(DateTime.Now).ToString() + ".xlsx"));
                }
            }

            if (!File.Exists(tempFile))
            {
                CreateSpreadSheet(tempFile, config);
            }
            XSSFWorkbook workbook;

            while (IsFileLocked(tempFile))
            {
                logger.LogWarning("[" + tempFile + "]: File is locked and cannot be written to. Trying again in 30 Seconds: {time}", DateTime.Now);
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

                AddEntry(workbook, sensor);
            }

            using (FileStream stream = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
                stream.Close();
            }

            CopyDataToPath(logger, tempFile, config.Destination);
        }

        public static void AddEntry(XSSFWorkbook workbook, Sensor sensor)
        {
            ISheet sheet = workbook.GetSheet(sensor.IP);

            IRow row = sheet.CreateRow(sheet.LastRowNum + 1);

            if (sensor.IsOffline)
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
            }
            else
            {
                row.CreateCell(0).SetCellValue(sensor.Temperature);
                row.CreateCell(1).SetCellValue(sensor.Humidity);
                row.CreateCell(2).SetCellValue(sensor.DewPoint);
                row.CreateCell(3).SetCellValue(sensor.IsRecording);
                row.CreateCell(4).SetCellValue(sensor.Date.ToString());
                row.CreateCell(5).SetCellValue(sensor.Time.ToString());
            }
        }

        private static void CreateWorkSheet(XSSFWorkbook workbook, Sensor sensor)
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

        private static void CreateSpreadSheet(string path, Config config)
        {
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                XSSFWorkbook workbook = new XSSFWorkbook();

                foreach (Sensor sensor in config.Sensors)
                {
                    CreateWorkSheet(workbook, sensor);
                }

                workbook.Write(stream);
                stream.Close();
            }
        }

        private static async void CopyDataToPath(ILogger<Worker> logger, string source, string destination)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            TimeOnly time = TimeOnly.FromDateTime(DateTime.Now);

            string file = Path.Combine(destination, "OSW Sensors.xlsx");

            await Task.Run(() => {
                if (File.Exists(file))
                {
                    while (IsFileLocked(file))
                    {
                        logger.LogWarning("[" + file + "]: File is locked and cannot be written to. Trying again in 30 Seconds: {time}", DateTime.Now);
                        Thread.Sleep(30000);

                        TimeSpan timeSpan = stopWatch.Elapsed;

                        if (timeSpan.TotalHours >= 1)
                        {
                            // "arvosgroup-com01c.mail.protection.outlook.com"
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

        private static bool IsFileLocked(string path)
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
    }
}
