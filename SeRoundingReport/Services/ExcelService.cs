using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

using NLog;

namespace SeRoundingReport.Services
{
    class ExcelService
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static bool GenerateReport(string fileNamePath, string reportName,DateTime reportDate, DataTable[] off, DataTable[] sup, DataTable[] offRawData, DataTable[] supRawData)
        {
            try
            {
                Color[] colors = new Color[] { Color.FromArgb(173,107,107), Color.FromArgb(138,74,74), Color.FromArgb(114,41,41)};

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage p = new ExcelPackage())
                {
                    p.Workbook.Properties.Author = "2020 Schneider Electric Homewood";
                    p.Workbook.Properties.Title = reportName;

                    ExcelWorksheet ws = CreateSheetWithDefaults(p, "Rounding Compliance");

                    // Merging cells and create a center heading for out table
                    ws.Cells[1, 1].Value = $"{reportName}";
                    ws.Cells[1, 1, 1, 7].Merge = true;
                    ws.Cells[1, 1, 1, 7].Style.Font.Bold = true;
                    ws.Cells[1, 1, 1, 7].Style.Font.Size = 14;
                    ws.Cells[1, 1, 1, 7].Style.Font.Color.SetColor(Color.Black);
                    ws.Cells[1, 1, 1, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[1, 1, 1, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Row(1).Height = 20;

                    ws.Cells["A3"].Value = "Date:";
                    ws.Cells["B3"].Value = $"{reportDate.AddDays(-6).ToShortDateString()} to {reportDate.ToShortDateString()}";

                    ws.Cells["A5"].Value = "Officer Building Rounds";
                    ws.Cells["A5"].Style.Font.Bold = true;
                    ws.Cells["A5"].Style.Font.Size = 14;

                    int rIndex = 5; // Last row modified

                    // Officer Rounding
                    int index = 0;
                    foreach (var t in off)
                    {
                        rIndex += 2;

                        var shift = index == 0 ? "1st Shift" : (index == 1 ? "2nd Shift" : "3rd Shift");
                        CreateShiftHeader(ws, rIndex, shift, colors[index]);
                        
                        rIndex++;
                        CreateHeader(ws, ref rIndex, t, Color.White);
                        CreateData(ws, ref rIndex, t);

                        index++;
                    }
                    rIndex += 2;
                    CreateTotalRow(ws, rIndex, Color.LightGray, off[0].Rows.Count);

                    // Supervisor Rounding
                    logger.Info($"Supervisor Post Text at {rIndex}");
                    rIndex += 2;
                    ws.Cells[rIndex, 1].Value = "Supervisor Post Rounds";
                    ws.Cells[rIndex, 1].Style.Font.Bold = true;
                    ws.Cells[rIndex, 1].Style.Font.Size = 14;

                    index = 0;
                    foreach (var t in sup)
                    {
                        rIndex += 2;

                        var shift = index == 0 ? "1st Shift" : (index == 1 ? "2nd Shift" : "3rd Shift");
                        CreateShiftHeader(ws, rIndex, shift, colors[index]);

                        rIndex++;
                        CreateHeader(ws, ref rIndex, t, Color.White);
                        CreateData(ws, ref rIndex, t, false);

                        index++;
                    }
                    rIndex += 2;
                    CreateTotalRow(ws, rIndex, Color.LightGray, sup[0].Rows.Count, false);

                    rIndex += 3;
                    CreateOverallTotalRow(ws, rIndex, Color.LightGray, sup[0].Rows.Count);

                    // Auto fit cells
                    for (int i = 1; i <= off[0].Columns.Count; i++)
                    {
                        ws.Column(i).AutoFit();
                    }

                    // Save raw data
                    index = 0;
                    rIndex = 0;
                    ExcelWorksheet ws2 = CreateSheetWithDefaults(p, "Supervisor Raw Data");
                    foreach (var s in supRawData)
                    {
                        if (rIndex > 0)
                            rIndex += 2;
                        else
                            rIndex++;

                        ws2.Cells[rIndex, 1].Value = index == 0 ? "1st Shift" : (index == 1 ? "2nd Shift" : "3rd Shift");
                        ws2.Cells[rIndex, 1].Style.Font.Bold = true;
                        ws2.Cells[rIndex, 1].Style.Font.Size = 16;
                        rIndex += 2;
                        CreateHeader(ws2, ref rIndex, s, Color.LightGray);
                        CreateRawData(ws2, ref rIndex, s);
                        index++;
                    }

                    for (int i = 1; i <= supRawData[0].Columns.Count; i++)
                    {
                        ws2.Column(i).AutoFit();
                    }

                    index = 0;
                    rIndex = 0;
                    ExcelWorksheet ws3 = CreateSheetWithDefaults(p, "Officer Raw Data");
                    foreach (var o in offRawData)
                    {
                        if (rIndex > 0)
                            rIndex += 2;
                        else
                            rIndex++;

                        ws3.Cells[rIndex, 1].Value = index == 0 ? "1st Shift" : (index == 1 ? "2nd Shift" : "3rd Shift");
                        ws3.Cells[rIndex, 1].Style.Font.Bold = true;
                        ws3.Cells[rIndex, 1].Style.Font.Size = 16;
                        rIndex += 2;
                        CreateHeader(ws3, ref rIndex, o, Color.LightGray);
                        CreateRawData(ws3, ref rIndex, o);
                        index++;
                    }

                    for (int i = 1; i <= offRawData[0].Columns.Count; i++)
                    {
                        ws3.Column(i).AutoFit();
                    }

                    // Save the report
                    string file = fileNamePath + $" {DateTime.Now.ToString("yyyyMMdd")}.xlsx";

                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(file, bin);

                    return true;
                }
            }
            catch (Exception ex) { logger.Error(ex, "ExcelService <GenerateReport> method."); Console.WriteLine($"ExcelService: {ex.Message}"); return false; }
        }

        #region Methods
        private static ExcelWorksheet CreateSheetWithDefaults(ExcelPackage p, string sheetName)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[sheetName];
            ws.Name = sheetName;
            // Worksheet defaults
            ws.Cells.Style.Font.Size = 11;
            ws.Cells.Style.Font.Name = "Calibri";
            ws.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
            ws.View.ZoomScale = 80;

            return ws;
        }

        private static void CreateHeader(ExcelWorksheet ws, ref int rowIndex, DataTable dt, Color? reportColor = null)
        {
            int colIndex = 1;
            foreach (DataColumn dc in dt.Columns) //Creating Headings
            {
                var cell = ws.Cells[rowIndex, colIndex];

                cell.Value = dc.ColumnName.Contains("Column") ? "" : dc.ColumnName;

                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.Black);
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(reportColor ?? Color.FromArgb(99, 89, 158));

                if (colIndex > 1)
                {
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                }

                colIndex++;
            }
        }

        private static void CreateData(ExcelWorksheet ws, ref int rowIndex, DataTable dt, bool isOfficer = true)
        {
            int colIndex = 0;
            foreach (DataRow dr in dt.Rows)
            {
                colIndex = 1;
                rowIndex++;

                foreach (DataColumn dc in dt.Columns)
                {
                    ws.Cells[rowIndex, colIndex].Value = dr[dc.ColumnName];

                    if (colIndex > 1)
                    {
                        ws.Cells[rowIndex, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[rowIndex, colIndex].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                    }

                    if (dc.ColumnName == "Total Required")
                    {
                        ws.Cells[rowIndex, colIndex].Formula = $"=(B{rowIndex})*(C{rowIndex})*7";
                    }
                    else if (dc.ColumnName == "Compliance")
                    {
                        ws.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#0\\.00%";
                        ws.Cells[rowIndex, colIndex].Formula = $"=(E{rowIndex})/(D{rowIndex}) * 100";
                    }
                    else if (dc.ColumnName == "Target")
                    {
                        ws.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#0\\.00%";
                    }

                    colIndex++;
                }
            }

            rowIndex++;
            for (int i = 1; i <= 7; i++)
            {
                if (i == 1)
                {
                    ws.Cells[rowIndex, i].Value = "Sub-Total";
                }
                if (i > 1 && i < 6)
                {
                    ws.Cells[rowIndex, i].Formula = $"=SUM({ws.Cells[rowIndex - dt.Rows.Count, i].Address}:{ws.Cells[rowIndex - 1, i].Address})";
                }
                if (i == 6)
                {
                    ws.Cells[rowIndex, i].Style.Numberformat.Format = "#0\\.00%";
                    ws.Cells[rowIndex, i].Formula = $"=(E{rowIndex})/(D{rowIndex}) * 100";
                }
                if (i == 7)
                {
                    ws.Cells[rowIndex, i].Style.Numberformat.Format = "#0\\.00%";
                    ws.Cells[rowIndex, i].Value = 90;
                }

                ws.Cells[rowIndex, i].Style.Font.Bold = true;

                if (i > 1)
                {
                    ws.Cells[rowIndex, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[rowIndex, i].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                }
            }
        }

        private static void CreateShiftHeader(ExcelWorksheet ws, int rIndex, string shift, Color headerColor)
        {
            ws.Cells[rIndex, 1].Value = shift;
            ws.Cells[rIndex, 1, rIndex, 7].Merge = true;
            ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Bold = true;
            ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Size = 14;
            ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Color.SetColor(Color.White);
            ws.Cells[rIndex, 1, rIndex, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[rIndex, 1, rIndex, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells[rIndex, 1, rIndex, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[rIndex, 1, rIndex, 7].Style.Fill.BackgroundColor.SetColor(headerColor);
        }

        private static void CreateTotalRow(ExcelWorksheet ws, int rIndex, Color rowColor, int doorCount, bool isOfficer = true)
        {
            int col1, col2, col3;
            col1 = doorCount * 2 + 10;
            col2 = doorCount + 6;
            col3 = 2;

            for (int i = 1; i <= 7; i++)
            {
                if (i == 1)
                {
                    ws.Cells[rIndex, i].Value = "Total";
                }
                if (i > 1 && i < 6)
                {
                    ws.Cells[rIndex, i].Formula = $"=SUM({ws.Cells[rIndex - col1, i].Address}+{ws.Cells[rIndex - col2, i].Address}+{ws.Cells[rIndex - col3, i].Address})";
                }
                if (i == 6)
                {
                    ws.Cells[rIndex, i].Style.Numberformat.Format = "#0\\.00%";
                    ws.Cells[rIndex, i].Formula = $"=(E{rIndex})/(D{rIndex}) * 100";
                }
                if (i == 7)
                {
                    ws.Cells[rIndex, i].Style.Numberformat.Format = "#0\\.00%";
                    ws.Cells[rIndex, i].Value = 90;
                }

                ws.Cells[rIndex, i].Style.Font.Bold = true;
                ws.Cells[rIndex, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[rIndex, i].Style.Fill.BackgroundColor.SetColor(rowColor);

                if (i > 1)
                {
                    ws.Cells[rIndex, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[rIndex, i].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                }
            }
        }

        private static void CreateOverallTotalRow(ExcelWorksheet ws, int rIndex, Color rowColor, int doorCount)
        {
            int col1, col2;
            col1 = doorCount * 3 + 19;
            col2 = 3;

            for (int i = 1; i <= 7; i++)
            {
                if (i == 1)
                {
                    ws.Cells[rIndex, i].Value = "Overall Total";
                }
                if (i > 1 && i < 6)
                {
                    ws.Cells[rIndex, i].Formula = $"=SUM({ws.Cells[rIndex - col1, i].Address}+{ws.Cells[rIndex - col2, i].Address})";
                }
                if (i == 6)
                {
                    ws.Cells[rIndex, i].Style.Numberformat.Format = "#0\\.00%";
                    ws.Cells[rIndex, i].Formula = $"=(E{rIndex})/(D{rIndex}) * 100";
                }
                if (i == 7)
                {
                    ws.Cells[rIndex, i].Style.Numberformat.Format = "#0\\.00%";
                    ws.Cells[rIndex, i].Value = 90;
                }

                ws.Cells[rIndex, i].Style.Font.Bold = true;
                ws.Cells[rIndex, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[rIndex, i].Style.Fill.BackgroundColor.SetColor(rowColor);

                if (i > 1)
                {
                    ws.Cells[rIndex, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[rIndex, i].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                }
            }
        }

        private static void CreateRawData(ExcelWorksheet ws, ref int rowIndex, DataTable dt, bool isOfficer = true)
        {
            int colIndex = 0;
            foreach (DataRow dr in dt.Rows)
            {
                colIndex = 1;
                rowIndex++;

                foreach (DataColumn dc in dt.Columns)
                {
                    ws.Cells[rowIndex, colIndex].Value = dr[dc.ColumnName];

                    colIndex++;
                }
            }
        }
        #endregion

    }
}
