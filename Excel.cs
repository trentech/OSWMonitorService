﻿using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Serilog;
using System.Diagnostics;
using System.Net.Mail;

namespace OSWMontiorService
{
    public class Excel
    {
        Config config;

        public Excel(Config config)
        {
            this.config = config;
        }

        public void AddAll()
        {
            string tempFile = Path.Combine(Config.PATH, "temp.xlsx");
            string dbFile = Path.Combine(config.DataType.Path, config.DataType.Name) + ".xlsx";

            if (!File.Exists(dbFile))
            {
                if (File.Exists(tempFile))
                {
                    Console.WriteLine(Path.Combine(Config.PATH, "temp" + "_" + DateTime.Now.ToString("MM-dd-yyyy_h:mm_tt") + ".xlsx"));
                    File.Move(tempFile, Path.Combine(Config.PATH, "temp" + "_" + DateTime.Now.ToString("MM-dd-yyyy_h-mm_tt") + ".xlsx"));
                }
            }

            if (!File.Exists(tempFile))
            {
                CreateSpreadSheet(tempFile, config);
            }
            XSSFWorkbook workbook;

            while (IsFileLocked(tempFile))
            {
                Log.Warning("[" + tempFile + "]: File is locked and cannot be written to. Trying again in 30 Seconds.");
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

            CopyDataToPath(tempFile, dbFile);
        }

        public void AddEntry(XSSFWorkbook workbook, Sensor sensor)
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

        private void CreateWorkSheet(XSSFWorkbook workbook, Sensor sensor)
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

        private void CreateSpreadSheet(string path, Config config)
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

        private async void CopyDataToPath(string source, string destination)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            TimeOnly time = TimeOnly.FromDateTime(DateTime.Now);

            await Task.Run(() => {
                if (File.Exists(destination))
                {
                    double index = 1;
                    while (IsFileLocked(destination))
                    {
                        double delay = config.DevMode ? 6 : config.Delay;
                        double check = (delay / 6) * index;

                        TimeSpan timeSpan = stopWatch.Elapsed;

                        if (timeSpan.TotalSeconds >= (delay * 60))
                        {
                            Log.Error("[" + destination + "]: File has been locked for " + Math.Round(timeSpan.TotalMinutes, 2) + " Minutes and cannot be written to. Timed out.");

                            if (!config.DevMode)
                            {
                                Mail mail = new Mail();

                                SmtpClient smtpClient = new SmtpClient(mail.STMP)
                                {
                                    Port = mail.Port,
                                    EnableSsl = mail.SSL,
                                };

                                MailMessage message = new MailMessage();
                                message.From = new MailAddress(mail.From);

                                foreach (string address in mail.Recipients)
                                {
                                    message.To.Add(new MailAddress(address));
                                }

                                message.Subject = "OSW Sensor File Locked. Timed Out";
                                message.Body = "[" + destination + "]: File has been locked for " + Math.Round(timeSpan.TotalMinutes, 2) + " Minutes and cannot be written to. Timed out.";

                                smtpClient.Send(message);
                            }

                            stopWatch.Stop();
                            return;
                        }
                        else if (timeSpan.TotalSeconds >= (check * 60))
                        {
                            index++;
                            Log.Warning("[" + destination + "]: File has been locked for " + Math.Round(timeSpan.TotalMinutes, 2) + " Minutes and cannot be written to. Sending Email.");

                            if (!config.DevMode)
                            {
                                Mail mail = new Mail();

                                SmtpClient smtpClient = new SmtpClient(mail.STMP)
                                {
                                    Port = mail.Port,
                                    EnableSsl = mail.SSL,
                                };

                                MailMessage message = new MailMessage();
                                message.From = new MailAddress(mail.From);

                                foreach (string address in mail.Recipients)
                                {
                                    message.To.Add(new MailAddress(address));
                                }

                                message.Subject = "OSW Sensor File Locked";
                                message.Body = "[" + destination + "] File has been locked for " + Math.Round(timeSpan.TotalMinutes, 2) + " Minutes and cannot be written to.";

                                smtpClient.Send(message);
                            }
                        }
                        else
                        {
                            Log.Warning("[" + destination + "]: File is locked and cannot be written to.");
                        }

                        Thread.Sleep(config.DevMode ? 10000 : 30000);
                    }

                    File.Delete(destination);
                }

                Log.Information("[" + destination + "]: Saving.");
                File.Copy(source, destination);
            });
        }

        private bool IsFileLocked(string path)
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
