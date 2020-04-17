﻿using System;
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
        public static bool GenerateReport(string fileNamePath, string reportName, DataTable[] dt, Color? reportColor = null)
        {
            try
            {
                Color[] colors = new Color[] { Color.FromArgb(173,107,107), Color.FromArgb(138,74,74), Color.FromArgb(114,41,41)};

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage p = new ExcelPackage())
                {
                    p.Workbook.Properties.Author = "2020 Schneider Electric Homewood";
                    p.Workbook.Properties.Title = reportName;

                    ExcelWorksheet ws = CreateSheetWithDefaults(p, reportName);

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
                    ws.Cells["B3"].Value = $"{DateTime.Now.AddDays(-7).ToShortDateString()} to {DateTime.Now.ToShortDateString()}";

                    ws.Cells["A5"].Value = "Officer Building Rounds";
                    ws.Cells["A5"].Style.Font.Bold = true;

                    int rIndex = 5; // Last row modified

                    int index = 0;
                    foreach (var t in dt)
                    {
                        rIndex += 2;

                        var shift = index == 0 ? "1st Shift" : (index == 1 ? "2nd Shift" : "3rd Shift");
                        CreateShiftHeader(ws, rIndex, shift, colors[index]);
                        
                        rIndex++;
                        CreateHeader(ws, ref rIndex, t, Color.White);
                        CreateData(ws, ref rIndex, t);

                        index++;
                    }

                    ////ws.Cells[7, 1].Value = $"1st Shift";
                    ////ws.Cells[7, 1, 7, 7].Merge = true;
                    ////ws.Cells[7, 1, 7, 7].Style.Font.Bold = true;
                    ////ws.Cells[7, 1, 7, 7].Style.Font.Size = 14;
                    ////ws.Cells[7, 1, 7, 7].Style.Font.Color.SetColor(Color.Black);
                    ////ws.Cells[7, 1, 7, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ////ws.Cells[7, 1, 7, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ////ws.Cells[7, 1, 7, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////ws.Cells[7, 1, 7, 7].Style.Fill.BackgroundColor.SetColor(reportColor ?? Color.FromArgb(183, 222, 232));

                    ////rIndex = 8;
                    ////CreateHeader(ws, ref rIndex, dt[0], reportColor ?? Color.White);
                    ////CreateData(ws, ref rIndex, dt[0]);

                    ////rIndex += 2;

                    ////ws.Cells[rIndex, 1].Value = $"2nd Shift";
                    ////ws.Cells[rIndex, 1, rIndex, 7].Merge = true;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Bold = true;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Size = 14;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Color.SetColor(Color.Black);
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Fill.BackgroundColor.SetColor(reportColor ?? Color.FromArgb(183, 222, 232));

                    ////rIndex++;
                    ////CreateHeader(ws, ref rIndex, dt[1], reportColor ?? Color.White);
                    ////CreateData(ws, ref rIndex, dt[1]);

                    ////rIndex += 2;

                    ////ws.Cells[rIndex, 1].Value = $"3rd Shift";
                    ////ws.Cells[rIndex, 1, rIndex, 7].Merge = true;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Bold = true;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Size = 14;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Font.Color.SetColor(Color.Black);
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////ws.Cells[rIndex, 1, rIndex, 7].Style.Fill.BackgroundColor.SetColor(reportColor ?? Color.FromArgb(183, 222, 232));

                    ////rIndex++;
                    ////CreateHeader(ws, ref rIndex, dt[2], reportColor ?? Color.White);
                    ////CreateData(ws, ref rIndex, dt[2]);

                    for (int i = 1; i <= dt[0].Columns.Count; i++)
                    {
                            ws.Column(i).AutoFit();
                    }

                    ////var chart = (ExcelBarChart)ws.Drawings.AddChart("crtHugsAlarms", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered);

                    ////chart.SetPosition(rowIndex + 2, 0, 2, 0);
                    ////chart.SetSize(400, 400);
                    ////chart.Series.Add("B8:B13", "A8:A13");
                    ////chart.Title.Text = "Hugs Alarms";
                    ////chart.Style = eChartStyle.Style31;
                    ////chart.Legend.Remove();

                    string file = fileNamePath + $" {DateTime.Now.ToString("yyyyMMdd")}.xlsx";
                    //string file = fileNamePath + $".xlsx";

                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(file, bin);

                    return true;
                }
            }
            catch (Exception ex) { logger.Error(ex, "ExcelService <GenerateReport> method."); Console.WriteLine($"ExcelService: {ex.Message}"); return false; }
        }

        #region local methods
        private static ExcelWorksheet CreateSheetWithDefaults(ExcelPackage p, string sheetName)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[0];
            ws.Name = "Rounding Compliance"; //sheetName; 
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

        private static void CreateData(ExcelWorksheet ws, ref int rowIndex, DataTable dt)
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
                        ws.Cells[rowIndex, colIndex].Formula = $"=(E{rowIndex})/(D{rowIndex})";
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
                    ws.Cells[rowIndex, i].Formula = $"=SUM({ws.Cells[rowIndex - 3, i].Address}:{ws.Cells[rowIndex - 1, i].Address})";
                }
                if (i == 6)
                {
                    ws.Cells[rowIndex, i].Style.Numberformat.Format = "#0\\.00%";
                    ws.Cells[rowIndex, i].Formula = $"=(E{rowIndex})/(D{rowIndex})";
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
        #endregion


    }
}