using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using OSWMonitorService.JSON;
using Serilog;
using System.Diagnostics;

namespace OSWMonitorService
{
    public class Excel
    {
        XSSFWorkbook workbook;
        string path;

        public Excel(Config config)
        {
            path = Path.Combine(config.DataType.Path, config.DataType.Name) + ".xlsx";

            if (!File.Exists(path))
            {
                using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    XSSFWorkbook workbook = new XSSFWorkbook();

                    foreach (Sensor sensor in Sensor.All())
                    {
                        CreateWorkSheet(workbook, sensor);
                    }

                    workbook.Write(stream);
                    stream.Close();
                }
            }

            checkLocked();

            using (FileStream stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(stream);
                stream.Close();
            }
        }

        private void checkLocked()
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            while (IsLocked(path))
            {
                if (stopwatch.Elapsed.TotalSeconds > 300)
                {
                    stopwatch.Stop();
                    Log.Error("[" + path + "]: File is locked and cannot be written to.");
                    throw new Exception("[" + path + "]: File is locked and cannot be written to");
                }

                Log.Warning("[" + path + "]: File is locked and cannot be written to. Trying again in 10 Seconds.");
                Thread.Sleep(10000);
            }

            stopwatch.Stop();
        }

        public void AddEntry(Sensor sensor)
        {
            ISheet sheet = workbook.GetSheet(sensor.IP);

            if (sheet == null)
            {
                sheet = CreateWorkSheet(workbook, sensor);
            }

            IRow row = sheet.CreateRow(sheet.LastRowNum + 1);

            ICell tempCell = row.CreateCell(0);
            ICell humidCell = row.CreateCell(1);
            ICell dewCell = row.CreateCell(2);
            ICell onlineCell = row.CreateCell(3);
            ICell dateCell = row.CreateCell(4);

            if (sheet.LastRowNum == 2)
            {
                ICellStyle amountStyle = workbook.CreateCellStyle();
                amountStyle.Alignment = HorizontalAlignment.Left;
                amountStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                tempCell.CellStyle = amountStyle;
                dewCell.CellStyle = amountStyle;
                sheet.SetDefaultColumnStyle(0, amountStyle);
                sheet.SetDefaultColumnStyle(1, amountStyle);

                ICellStyle percentStyle = workbook.CreateCellStyle();
                percentStyle.Alignment = HorizontalAlignment.Left;
                percentStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
                humidCell.CellStyle = percentStyle;
                sheet.SetDefaultColumnStyle(2, percentStyle);

                ICellStyle stringStyle = workbook.CreateCellStyle();
                stringStyle.Alignment = HorizontalAlignment.Left;
                stringStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("TEXT");
                onlineCell.CellStyle = stringStyle;
                sheet.SetDefaultColumnStyle(3, stringStyle);

                ICellStyle dateStyle = workbook.CreateCellStyle();
                dateStyle.Alignment = HorizontalAlignment.Left;
                dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("MM/dd/yyyy hh:mm AM/PM");
                dateCell.CellStyle = dateStyle;
                sheet.SetDefaultColumnStyle(4, dateStyle);
            }
            else
            {
                tempCell.CellStyle = sheet.GetRow(sheet.LastRowNum - 1).GetCell(0).CellStyle;
                humidCell.CellStyle = sheet.GetRow(sheet.LastRowNum - 1).GetCell(1).CellStyle;
                dewCell.CellStyle = sheet.GetRow(sheet.LastRowNum - 1).GetCell(2).CellStyle;
                onlineCell.CellStyle = sheet.GetRow(sheet.LastRowNum - 1).GetCell(3).CellStyle;
                dateCell.CellStyle = sheet.GetRow(sheet.LastRowNum - 1).GetCell(4).CellStyle;
            }

            tempCell.SetCellValue(sensor.Temperature);
            humidCell.SetCellValue(sensor.Humidity / 100);
            dewCell.SetCellValue(sensor.DewPoint);
            onlineCell.SetCellValue(sensor.IsOnline);
            dateCell.SetCellValue(sensor.DateTime);

            checkLocked();

            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
                stream.Close();
            }
        }

        private ISheet CreateWorkSheet(XSSFWorkbook workbook, Sensor sensor)
        {
            ISheet sheet = workbook.CreateSheet(sensor.IP);
            sheet.DefaultColumnWidth = 20;

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
            cell.SetCellValue("Online");

            cell = row.CreateCell(4);
            cell.CellStyle = style;
            cell.SetCellValue("DateTime");

            return sheet;
        }

        private bool IsLocked(string path)
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