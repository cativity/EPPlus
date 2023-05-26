﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Text;
using System.Globalization;
using System.Threading.Tasks;
using OfficeOpenXml.Export.HtmlExport.Interfaces;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.Table;

namespace EPPlusTest.Export.HtmlExport
{
    [TestClass]
    public class TableExporterTests : TestBase
    {
#if !NET35 && !NET40
        [TestMethod]
        public void ShouldExportHeadersAsync()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Age";
            sheet.Cells["A2"].Value = "John Doe";
            sheet.Cells["B2"].Value = "23";
            ExcelTable? table = sheet.Tables.Add(sheet.Cells["A1:B2"], "myTable");
            table.TableStyle = TableStyles.Dark1;
            table.ShowHeader = true;
            using MemoryStream? ms = new MemoryStream();
            IExcelHtmlTableExporter? exporter = table.CreateHtmlExporter();
            exporter.RenderHtmlAsync(ms).Wait();
            StreamReader? sr = new StreamReader(ms);
            ms.Position = 0;
            string? result = sr.ReadToEnd();
        }
#endif

        [TestMethod]
        public void ShouldExportHeadersWithNoAccessibilityAttributes()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Age";
            sheet.Cells["A2"].Value = "John Doe";
            sheet.Cells["B2"].Value = 23;
            ExcelTable? table = sheet.Tables.Add(sheet.Cells["A1:B2"], "myTable");
            table.TableStyle = TableStyles.Dark1;
            table.ShowHeader = true;
            IExcelHtmlTableExporter? exporter = table.CreateHtmlExporter();
            exporter.Settings.Configure(x =>
            {
                x.TableId = "myTable";
                x.Minify = true;
                x.Accessibility.TableSettings.AddAccessibilityAttributes = false;
            });
            string? html = exporter.GetHtmlString();
            using MemoryStream? ms = new MemoryStream();
            exporter.RenderHtml(ms);
            StreamReader? sr = new StreamReader(ms);
            ms.Position = 0;
            string? result = sr.ReadToEnd();
            string? expectedHtml = "<table class=\"epplus-table ts-dark1 ts-dark1-header ts-dark1-row-stripes\" id=\"myTable\"><thead><tr><th data-datatype=\"string\" class=\"epp-al\">Name</th><th data-datatype=\"number\" class=\"epp-al\">Age</th></tr></thead><tbody><tr><td>John Doe</td><td data-value=\"23\" class=\"epp-ar\">23</td></tr></tbody></table>";
            Assert.AreEqual(expectedHtml, result);
        }
        [TestMethod]
        public void ShouldExportHeadersWithAccessibilityAttributes()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Age";
            sheet.Cells["A2"].Value = "John Doe";
            sheet.Cells["B2"].Value = 23;
            ExcelTable? table = sheet.Tables.Add(sheet.Cells["A1:B2"], "myTable");
            table.TableStyle = TableStyles.Dark1;
            table.ShowHeader = true;

            IExcelHtmlTableExporter? exporter = table.CreateHtmlExporter();
            exporter.Settings.Configure(x =>
            {
                x.TableId = "myTable";
                x.Minify = true;
            });

            using MemoryStream? ms = new MemoryStream();
            exporter.RenderHtml(ms);
            StreamReader? sr = new StreamReader(ms);
            ms.Position = 0;
            string? result = sr.ReadToEnd();
            string? expectedHtml = "<table class=\"epplus-table ts-dark1 ts-dark1-header ts-dark1-row-stripes\" id=\"myTable\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\" role=\"columnheader\" scope=\"col\">Name</th><th data-datatype=\"number\" class=\"epp-al\" role=\"columnheader\" scope=\"col\">Age</th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\">John Doe</td><td data-value=\"23\" role=\"cell\" class=\"epp-ar\">23</td></tr></tbody></table>";
            Assert.AreEqual(expectedHtml, result);
        }

        [TestMethod]
        public void ExportAllTableStyles()
        {
            string path = _worksheetPath + "TableStyles";
            CreatePathIfNotExists(path);
            using ExcelPackage? p = OpenPackage("TableStylesToHtml.xlsx", true);
            foreach (TableStyles e in Enum.GetValues(typeof(TableStyles)))
            {
                if (!(e == TableStyles.Custom || e == TableStyles.None))
                {
                    ExcelWorksheet? ws = p.Workbook.Worksheets.Add(e.ToString());
                    LoadTestdata(ws);
                    ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{e}");
                    tbl.TableStyle = e;

                    IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
                    string? html = exporter.GetSinglePage();

                    File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                }
            }
            SaveAndCleanup(p);
        }
        [TestMethod]
        public async Task ExportAllTableStylesAsync()
        {
            string path = _worksheetPath + "TableStylesAsync";
            CreatePathIfNotExists(path);
            using ExcelPackage? p = OpenPackage("TableStylesToHtml.xlsx", true);
            foreach (TableStyles e in Enum.GetValues(typeof(TableStyles)))
            {
                if (!(e == TableStyles.Custom || e == TableStyles.None))
                {
                    ExcelWorksheet? ws = p.Workbook.Worksheets.Add(e.ToString());
                    LoadTestdata(ws);
                    ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{e}");
                    tbl.TableStyle = e;

                    IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
                    string? html = await exporter.GetSinglePageAsync();

                    File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                }
            }
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void ExportAllFirstLastTableStyles()
        {
            string path = _worksheetPath + "TableStylesFirstLast";
            CreatePathIfNotExists(path);
            using ExcelPackage? p = OpenPackage("TableStylesToHtmlFirstLastCol.xlsx", true);
            foreach (TableStyles e in Enum.GetValues(typeof(TableStyles)))
            {
                if (!(e == TableStyles.Custom || e == TableStyles.None))
                {
                    ExcelWorksheet? ws = p.Workbook.Worksheets.Add(e.ToString());
                    LoadTestdata(ws);
                    ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{e}");
                    tbl.ShowFirstColumn = true;
                    tbl.ShowLastColumn = true;
                    tbl.TableStyle = e;

                    IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
                    string? html = exporter.GetSinglePage();

                    File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                }
            }
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void ExportAllCustomTableStyles()
        {
            string path = _worksheetPath + "TableStylesCustomFills";
            CreatePathIfNotExists(path);
            using ExcelPackage? p = OpenPackage("TableStylesToHtmlPatternFill.xlsx", true);
            foreach (ExcelFillStyle fs in Enum.GetValues(typeof(ExcelFillStyle)))
            {
                ExcelWorksheet? ws = p.Workbook.Worksheets.Add($"PatterFill-{fs}");
                LoadTestdata(ws);
                ExcelTableNamedStyle? ts = p.Workbook.Styles.CreateTableStyle($"CustomPattern-{fs}", TableStyles.Medium9);
                ts.FirstRowStripe.Style.Fill.Style = eDxfFillStyle.PatternFill;
                ts.FirstRowStripe.Style.Fill.PatternType = fs;
                ts.FirstRowStripe.Style.Fill.PatternColor.Tint = 0.10;
                ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{fs}");
                tbl.StyleName = ts.Name;

                IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
                string? html = exporter.GetSinglePage();
                File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
            }
            SaveAndCleanup(p);
        }
        [TestMethod]
        public async Task ExportAllCustomTableStylesAsync()
        {
            string path = _worksheetPath + "TableStylesCustomFillsAsync";
            CreatePathIfNotExists(path);
            using ExcelPackage? p = OpenPackage("TableStylesToHtmlPatternFill.xlsx", true);
            foreach (ExcelFillStyle fs in Enum.GetValues(typeof(ExcelFillStyle)))
            {
                ExcelWorksheet? ws = p.Workbook.Worksheets.Add($"PatterFill-{fs}");
                LoadTestdata(ws);
                ExcelTableNamedStyle? ts = p.Workbook.Styles.CreateTableStyle($"CustomPattern-{fs}", TableStyles.Medium9);
                ts.FirstRowStripe.Style.Fill.Style = eDxfFillStyle.PatternFill;
                ts.FirstRowStripe.Style.Fill.PatternType = fs;
                ts.FirstRowStripe.Style.Fill.PatternColor.Tint = 0.10;
                ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{fs}");
                tbl.StyleName = ts.Name;

                IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
                string? html = await exporter.GetSinglePageAsync();
                File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
            }
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void ExportAllGradientTableStyles()
        {
            string path = _worksheetPath + "TableStylesGradientFills";
            CreatePathIfNotExists(path);
            using ExcelPackage? p = OpenPackage("TableStylesToHtmlGradientFill.xlsx", true);
            ExcelWorksheet? ws = p.Workbook.Worksheets.Add($"PatterFill-Gradient");
            LoadTestdata(ws);
            ExcelTableNamedStyle? ts = p.Workbook.Styles.CreateTableStyle($"CustomPattern-Gradient1", TableStyles.Medium9);
            ts.FirstRowStripe.Style.Fill.Style = eDxfFillStyle.GradientFill;
            ts.FirstRowStripe.Style.Fill.Gradient.GradientType = eDxfGradientFillType.Path;
            ExcelDxfGradientFillColor? c1 = ts.FirstRowStripe.Style.Fill.Gradient.Colors.Add(0);
            c1.Color.Color = Color.White;

            ExcelDxfGradientFillColor? c2 = ts.FirstRowStripe.Style.Fill.Gradient.Colors.Add(100);
            c2.Color.Color = Color.FromArgb(0x44, 0x72, 0xc4);

            ts.FirstRowStripe.Style.Fill.Gradient.Bottom = 0.5;
            ts.FirstRowStripe.Style.Fill.Gradient.Top = 0.5;
            ts.FirstRowStripe.Style.Fill.Gradient.Left = 0.5;
            ts.FirstRowStripe.Style.Fill.Gradient.Right = 0.5;

            ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tblGradient");
            tbl.StyleName = ts.Name;

            IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
            string? html = exporter.GetSinglePage();
            File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void ExportTableWithCellStylesStyles()
        {
            string path = _worksheetPath + "TableStylesCellStyles";
            CreatePathIfNotExists(path);
            using ExcelPackage? p = OpenPackage("TableStylesToHtmlCellStyles.xlsx", true);
            ExcelWorksheet? ws = p.Workbook.Worksheets.Add($"CellStyles");
            LoadTestdata(ws, 100, 1, 1, true);

            ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:E101"], $"tblGradient");
            tbl.TableStyle = TableStyles.Dark3;
            ws.Cells["A1"].Style.Font.Italic = true;
            ws.Cells["B1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells["C5"].Style.Font.Size = 18;
            tbl.Columns[0].TotalsRowLabel = "Total";
            IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
            string? html = exporter.GetSinglePage();
            File.WriteAllText($"{path}\\table-{tbl.StyleName}-CellStyle.html", html);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void ShouldExportWithOtherCultureInfo()
        {
            string path = _worksheetPath + "culture";
            CreatePathIfNotExists(path);
            using ExcelPackage? p = new ExcelPackage();
            ExcelWorksheet? ws = p.Workbook.Worksheets.Add($"CellStyles");
            LoadTestdata(ws, 100, 1, 1, true);

            ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:E101"], $"tblGradient");
            tbl.TableStyle = TableStyles.Dark3;
            ws.Cells["A1"].Style.Font.Italic = true;
            ws.Cells["B1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells["C5"].Style.Font.Size = 18;
            tbl.Columns[0].TotalsRowLabel = "Total";

            IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
            exporter.Settings.Culture = new CultureInfo("en-US");
            string? html = exporter.GetSinglePage();

            File.WriteAllText($"{path}\\table-{tbl.StyleName}-CellStyle.html", html);
        }
        [TestMethod]
        public void ValidateConfigureAndResetToDefault()
        {
            using ExcelPackage? p = new ExcelPackage();
            ExcelWorksheet? ws = p.Workbook.Worksheets.Add($"Sheet1");
            LoadTestdata(ws, 100, 1, 1, true);

            ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:E101"], $"tblGradient");
            tbl.TableStyle = TableStyles.Dark3;
            IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
            exporter.Settings.Configure(x =>
            {
                x.Encoding = Encoding.Unicode;
                x.Culture = new CultureInfo("en-GB");
                x.TableId = "Table1";
                x.RenderDataAttributes = false;
                x.Css.Exclude.TableStyle.Border = eBorderExclude.Right | eBorderExclude.Left;
                x.Css.Exclude.TableStyle.HorizontalAlignment = true;
                x.Css.Exclude.CellStyle.Fill = true;
                x.AdditionalTableClassNames.Add("ATC1");
                x.Accessibility.TableSettings.AriaLabel = "AriaLabel1";
                x.Accessibility.TableSettings.TableRole = "TableRoll1";
                x.Accessibility.TableSettings.AddAccessibilityAttributes = false;
            });

            HtmlTableExportSettings? s = exporter.Settings;
            Assert.AreEqual(Encoding.Unicode, s.Encoding);
            Assert.AreEqual("en-GB", s.Culture.Name);
            Assert.AreEqual("Table1", s.TableId);
            Assert.IsFalse(s.RenderDataAttributes);
            Assert.AreEqual(eBorderExclude.Right | eBorderExclude.Left, s.Css.Exclude.TableStyle.Border);
            Assert.IsTrue(s.Css.Exclude.TableStyle.HorizontalAlignment);
            Assert.IsTrue(s.Css.Exclude.CellStyle.Fill);
            Assert.AreEqual("ATC1", s.AdditionalTableClassNames[0]);
            Assert.AreEqual("AriaLabel1", s.Accessibility.TableSettings.AriaLabel);
            Assert.AreEqual("TableRoll1", s.Accessibility.TableSettings.TableRole);
            Assert.IsFalse(s.Accessibility.TableSettings.AddAccessibilityAttributes);

            exporter.Settings.ResetToDefault();

            s = exporter.Settings;
            Assert.AreEqual(Encoding.UTF8, s.Encoding);
            Assert.AreEqual(CultureInfo.CurrentCulture.Name, s.Culture.Name);
            Assert.IsTrue(string.IsNullOrEmpty(s.TableId));
            Assert.IsTrue(s.RenderDataAttributes);
            Assert.AreEqual(0, (int)s.Css.Exclude.TableStyle.Border);
            Assert.IsFalse(s.Css.Exclude.TableStyle.HorizontalAlignment);
            Assert.IsFalse(s.Css.Exclude.CellStyle.Fill);
            Assert.AreEqual(0, s.AdditionalTableClassNames.Count);
            Assert.IsTrue(string.IsNullOrEmpty(s.Accessibility.TableSettings.AriaLabel));
            Assert.AreEqual("table", s.Accessibility.TableSettings.TableRole);
            Assert.IsTrue(s.Accessibility.TableSettings.AddAccessibilityAttributes);
        }
        [TestMethod]
        public void ShouldExportRichTextAsInlineHtml()
        {
            using ExcelPackage? p = new ExcelPackage();
            ExcelWorksheet? ws = p.Workbook.Worksheets.Add($"RichText");

            ExcelRichTextCollection? rt = ws.Cells["A1"].RichText;
            ExcelRichText? rt1 = rt.Add("Header");
            rt1.Color = Color.Red;
            ExcelRichText? rt2 = rt.Add(" 1");
            rt2.Color = Color.Blue;

            rt = ws.Cells["B1"].RichText;
            rt1 = rt.Add("Header");
            rt1.Italic = true;
            rt1.Bold = true;
            rt2 = rt.Add(" 2");
            rt2.Strike = true;

            rt = ws.Cells["C1"].RichText;
            rt1 = rt.Add("Header");
            rt1.FontName = "Arial";
            rt1.Size = 12;
            rt2 = rt.Add(" 3");
            rt2.UnderLine = true;

            rt = ws.Cells["A2"].RichText;
            rt1 = rt.Add("Text");
            rt1.Color = Color.Green;
            rt2 = rt.Add(" 1");
            rt2.Color = Color.Yellow;

            rt = ws.Cells["B2"].RichText;
            rt1 = rt.Add("Text");
            rt1.Italic = true;
            rt1.Bold = true;
            rt2 = rt.Add(" 2");
            rt2.Strike = true;

            rt = ws.Cells["C2"].RichText;
            rt1 = rt.Add("Text");
            rt1.FontName = "Times New Roman";
            rt1.Size = 8;
            rt2 = rt.Add(" 3");
            rt2.UnderLine = true;


            ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:C2"], $"tblRichtext");
            tbl.TableStyle = TableStyles.Dark5;

            IExcelHtmlTableExporter? exporter = tbl.CreateHtmlExporter();
            string? html = exporter.GetHtmlString();
            string? htmlCss = exporter.GetSinglePage();
        }
        [TestMethod]
        public async Task WriteImages_TableAsync()
        {
            using ExcelPackage? p = OpenTemplatePackage("20-CreateAFileSystemReport-Table.xlsx");
            ExcelWorksheet? sheet = p.Workbook.Worksheets[0];
            IExcelHtmlTableExporter? exporter = sheet.Tables[0].CreateHtmlExporter();
            exporter.Settings.SetColumnWidth = true;
            exporter.Settings.SetRowHeight = true;
            exporter.Settings.Pictures.Include = ePictureInclude.Include;
            exporter.Settings.Minify = false;
            string? html = exporter.GetSinglePage();
            string? htmlAsync = await exporter.GetSinglePageAsync();
            File.WriteAllText("c:\\temp\\" + sheet.Name + "-table.html", html);
            File.WriteAllText("c:\\temp\\" + sheet.Name + "-table-async.html", htmlAsync);
            Assert.AreEqual(html, htmlAsync);
        }
        [TestMethod]
        public async Task WriteTableFromRange()
        {
            using ExcelPackage? p = OpenTemplatePackage("20-CreateAFileSystemReport.xlsx");
            ExcelWorksheet? sheet = p.Workbook.Worksheets[1];
            IExcelHtmlRangeExporter? exporterRange = sheet.Tables[0].Range.CreateHtmlExporter();
            exporterRange.Settings.SetColumnWidth = true;
            exporterRange.Settings.SetRowHeight = true;
            exporterRange.Settings.Minify = false;
            exporterRange.Settings.TableStyle = eHtmlRangeTableInclude.ClassNamesOnly;
            string? html = exporterRange.GetHtmlString();
            string? htmlAsync = await exporterRange.GetHtmlStringAsync();

            string? css = exporterRange.GetCssString();
            string? cssAsync = await exporterRange.GetCssStringAsync();

            string? outputHtml = string.Format("<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>", html, css);

            File.WriteAllText("c:\\temp\\TableRangeCombined.html", outputHtml);

            Assert.AreEqual(html, htmlAsync);
            Assert.AreEqual(css, cssAsync);
        }
        [TestMethod]
        public async Task WriteMultipleRangeWithTableAndRange()
        {
            using ExcelPackage? p = OpenTemplatePackage("20-CreateAFileSystemReport.xlsx");
            ExcelWorksheet? sheet1 = p.Workbook.Worksheets[0];
            ExcelWorksheet? sheet2 = p.Workbook.Worksheets[1];
            IExcelHtmlRangeExporter? exporterRange = p.Workbook.CreateHtmlExporter(
                                                                                   sheet2.Tables[0].Range,
                                                                                   sheet1.Cells["A1:E30"],
                                                                                   sheet2.Tables[2].Range,
                                                                                   sheet2.Tables[1].Range);
            exporterRange.Settings.SetColumnWidth = true;
            exporterRange.Settings.SetRowHeight = true;
            exporterRange.Settings.Minify = false;
            exporterRange.Settings.TableStyle = eHtmlRangeTableInclude.Include;
            exporterRange.Settings.Pictures.Include = ePictureInclude.Include;
            string? html1 = exporterRange.GetHtmlString(0);
            string? html2 = exporterRange.GetHtmlString(1);
            string? html3 = exporterRange.GetHtmlString(2);
            string? html4 = exporterRange.GetHtmlString(3);

            string? css = exporterRange.GetCssString();
            string? cssAsync = await exporterRange.GetCssStringAsync();

            string? outputHtml = string.Format("<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{4}</style></head>\r\n<body>\r\n{0}<hr>{1}<hr>{2}<hr>{3}<hr></body>\r\n</html>", html1, html2, html3, html4, css);

            File.WriteAllText("c:\\temp\\RangeAndThreeTables.html", outputHtml);

            Assert.AreEqual(css, cssAsync);
        }
    }
}
