/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.VBA;

namespace EPPlusTest;

/// <summary>
/// This class contains testcases for issues on Codeplex and Github.
/// All tests requiering an template should be set to ignored as it's not practical to include all xlsx templates in the project.
/// </summary>
[TestClass]
public class Issues : TestBase
{
    [ClassInitialize]
    public static void Init(TestContext context)
    {
    }

    [ClassCleanup]
    public static void Cleanup()
    {
    }

    [TestInitialize]
    public void Initialize()
    {
    }

    [TestMethod]
    public void Issue15041()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("Test");
        ws.Cells["A1"].Value = 202100083;
        ws.Cells["A1"].Style.Numberformat.Format = "00\\.00\\.00\\.000\\.0";
        Assert.AreEqual("02.02.10.008.3", ws.Cells["A1"].Text);
        ws.Dispose();
    }

    [TestMethod]
    public void Issue15031()
    {
        double d = OfficeOpenXml.Utils.ConvertUtil.GetValueDouble(new TimeSpan(35, 59, 1));
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("Test");
        ws.Cells["A1"].Value = d;
        ws.Cells["A1"].Style.Numberformat.Format = "[t]:mm:ss";
        ws.Dispose();
    }

    [TestMethod]
    public void Issue15022()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("Test");
        ws.Cells.AutoFitColumns();
        ws.Cells["A1"].Style.Numberformat.Format = "0";
        ws.Cells.AutoFitColumns();
    }

    [TestMethod]
    public void Issue15056()
    {
        using ExcelPackage? ep = OpenPackage(@"output.xlsx", true);
        ExcelWorksheet? s = ep.Workbook.Worksheets.Add("test");
        s.Cells["A1:A2"].Formula = ""; // or null, or non-empty whitespace, with same result
        ep.Save();
    }

    [TestMethod]
    public void Issue15113()
    {
        ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("t");
        ws.Cells["A1"].Value = " Performance Update";
        ws.Cells["A1:H1"].Merge = true;
        ws.Cells["A1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
        ws.Cells["A1:H1"].Style.Font.Size = 14;
        ws.Cells["A1:H1"].Style.Font.Color.SetColor(Color.Red);
        ws.Cells["A1:H1"].Style.Font.Bold = true;
        SaveWorkbook(@"merge.xlsx", p);
        p.Dispose();
    }

    [TestMethod]
    public void Issue15141()
    {
        using ExcelPackage package = new ExcelPackage();
        using ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Test");
        sheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
        sheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
        sheet.Cells[1, 1, 1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        sheet.Cells[1, 5, 2, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        _ = sheet.Column(3);
    }

    [TestMethod]
    public void Issue15123()
    {
        ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("t");
        using DataTable? dt = new DataTable();
        _ = dt.Columns.Add("String", typeof(string));
        _ = dt.Columns.Add("Int", typeof(int));
        _ = dt.Columns.Add("Bool", typeof(bool));
        _ = dt.Columns.Add("Double", typeof(double));
        _ = dt.Columns.Add("Date", typeof(DateTime));

        DataRow? dr = dt.NewRow();
        dr[0] = "Row1";
        dr[1] = 1;
        dr[2] = true;
        dr[3] = 1.5;
        dr[4] = new DateTime(2014, 12, 30);
        dt.Rows.Add(dr);

        dr = dt.NewRow();
        dr[0] = "Row2";
        dr[1] = 2;
        dr[2] = false;
        dr[3] = 2.25;
        dr[4] = new DateTime(2014, 12, 31);
        dt.Rows.Add(dr);

        _ = ws.Cells["A1"].LoadFromDataTable(dt, true);
        ws.Cells["D2:D3"].Style.Numberformat.Format = "(* #,##0.00);_(* (#,##0.00);_(* \"-\"??_);(@)";

        ws.Cells["E2:E3"].Style.Numberformat.Format = "mm/dd/yyyy";
        ws.Cells.AutoFitColumns();
        Assert.AreNotEqual(ws.Cells[2, 5].Text, "");
    }

    [TestMethod]
    public void Issue15128()
    {
        ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("t");
        ws.Cells["A1"].Value = 1;
        ws.Cells["B1"].Value = 2;
        ws.Cells["B2"].Formula = "A1+$B$1";
        ws.Cells["C1"].Value = "Test";
        ws.Cells["A1:B2"].Copy(ws.Cells["C1"]);
        ws.Cells["B2"].Copy(ws.Cells["D1"]);
        SaveWorkbook("Copy.xlsx", p);
        p.Dispose();
    }

    [TestMethod]
    public void IssueMergedCells()
    {
        ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("t");
        ws.Cells["A1:A5,C1:C8"].Merge = true;
        ws.Cells["C1:C8"].Merge = false;
        ws.Cells["A1:A8"].Merge = false;
        p.Dispose();
    }

    public class cls1
    {
        public int prop1 { get; set; }
    }

    public class cls2 : cls1
    {
        public string prop2 { get; set; }
    }

    [TestMethod]
    public void LoadFromColIssue()
    {
        List<cls1>? l = new List<cls1>();

        l.Add(new cls1() { prop1 = 1 });
        l.Add(new cls2() { prop1 = 1, prop2 = "test1" });

        ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Test");

        _ = ws.Cells["A1"]
          .LoadFromCollection(l,
                              true,
                              TableStyles.Light16,
                              BindingFlags.Instance | BindingFlags.Public,
                              new MemberInfo[] { typeof(cls2).GetProperty("prop2") });
    }

    [TestMethod]
    public void Issue15168()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Test");
        ws.Cells[1, 1].Value = "A1";
        ws.Cells[2, 1].Value = "A2";

        ws.Cells[2, 1].Value = ws.Cells[1, 1].Value;
        Assert.AreEqual("A1", ws.Cells[1, 1].Value);
    }

    [TestMethod]
    public void Issue15179()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("MergeDeleteBug");
        ws.Cells["E3:F3"].Merge = true;
        ws.Cells["E3:F3"].Merge = false;
        ws.DeleteRow(2, 6);
        ws.Cells["A1"].Value = 0;
    }

    [TestMethod]
    public void Issue15212()
    {
        string? s = "_(\"R$ \"* #,##0.00_);_(\"R$ \"* (#,##0.00);_(\"R$ \"* \"-\"??_);_(@_) )";
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("StyleBug");
        ws.Cells["A1"].Value = 5698633.64;
        ws.Cells["A1"].Style.Numberformat.Format = s;
    }

    [TestMethod]
    /**** Pivottable issue ****/
    public void Issue()
    {
        using ExcelPackage? p = OpenPackage("pivottable.xlsx", true);
        LoadData(p);
        BuildPivotTable1(p);
        BuildPivotTable2(p);
        p.Save();
    }

    private static void LoadData(ExcelPackage p)
    {
        // add a new worksheet to the empty workbook
        ExcelWorksheet wsData = p.Workbook.Worksheets.Add("Data");

        //Add the headers
        wsData.Cells[1, 1].Value = "INVOICE_DATE";
        wsData.Cells[1, 2].Value = "TOTAL_INVOICE_PRICE";
        wsData.Cells[1, 3].Value = "EXTENDED_PRICE_VARIANCE";
        wsData.Cells[1, 4].Value = "AUDIT_LINE_STATUS";
        wsData.Cells[1, 5].Value = "RESOLUTION_STATUS";
        wsData.Cells[1, 6].Value = "COUNT";

        //Add some items...
        wsData.Cells["A2"].Value = Convert.ToDateTime("04/2/2012");
        wsData.Cells["B2"].Value = 33.63;
        wsData.Cells["C2"].Value = -.87;
        wsData.Cells["D2"].Value = "Unfavorable Price Variance";
        wsData.Cells["E2"].Value = "Pending";
        wsData.Cells["F2"].Value = 1;

        wsData.Cells["A3"].Value = Convert.ToDateTime("04/2/2012");
        wsData.Cells["B3"].Value = 43.14;
        wsData.Cells["C3"].Value = -1.29;
        wsData.Cells["D3"].Value = "Unfavorable Price Variance";
        wsData.Cells["E3"].Value = "Pending";
        wsData.Cells["F3"].Value = 1;

        wsData.Cells["A4"].Value = Convert.ToDateTime("11/8/2011");
        wsData.Cells["B4"].Value = 55;
        wsData.Cells["C4"].Value = -2.87;
        wsData.Cells["D4"].Value = "Unfavorable Price Variance";
        wsData.Cells["E4"].Value = "Pending";
        wsData.Cells["F4"].Value = 1;

        wsData.Cells["A5"].Value = Convert.ToDateTime("11/8/2011");
        wsData.Cells["B5"].Value = 38.72;
        wsData.Cells["C5"].Value = -5.00;
        wsData.Cells["D5"].Value = "Unfavorable Price Variance";
        wsData.Cells["E5"].Value = "Pending";
        wsData.Cells["F5"].Value = 1;

        wsData.Cells["A6"].Value = Convert.ToDateTime("3/4/2011");
        wsData.Cells["B6"].Value = 77.44;
        wsData.Cells["C6"].Value = -1.55;
        wsData.Cells["D6"].Value = "Unfavorable Price Variance";
        wsData.Cells["E6"].Value = "Pending";
        wsData.Cells["F6"].Value = 1;

        wsData.Cells["A7"].Value = Convert.ToDateTime("3/4/2011");
        wsData.Cells["B7"].Value = 127.55;
        wsData.Cells["C7"].Value = -10.50;
        wsData.Cells["D7"].Value = "Unfavorable Price Variance";
        wsData.Cells["E7"].Value = "Pending";
        wsData.Cells["F7"].Value = 1;

        using (ExcelRange? range = wsData.Cells[2, 1, 7, 1])
        {
            range.Style.Numberformat.Format = "mm-dd-yy";
        }

        wsData.Cells.AutoFitColumns(0);
    }

    private static void BuildPivotTable1(ExcelPackage p)
    {
        ExcelWorksheet? wsData = p.Workbook.Worksheets["Data"];
        string? totalRows = wsData.Dimension.Address;
        ExcelRange data = wsData.Cells[totalRows];

        ExcelWorksheet? wsAuditPivot = p.Workbook.Worksheets.Add("Pivot1");

        ExcelPivotTable? pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit1");
        pivotTable1.ColumnGrandTotals = true;
        ExcelPivotTableField? rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);

        rowField.AddDateGrouping(eDateGroupBy.Years);
        ExcelPivotTableField? yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
        yearField.Name = "Year";

        _ = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

        ExcelPivotTableDataField? TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
        TotalSpend.Name = "Total Spend";
        TotalSpend.Format = "$##,##0";

        ExcelPivotTableDataField? CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
        CountInvoicePrice.Name = "Total Lines";
        CountInvoicePrice.Format = "##,##0";

        pivotTable1.DataOnRows = false;
    }

    private static void BuildPivotTable2(ExcelPackage p)
    {
        ExcelWorksheet? wsData = p.Workbook.Worksheets["Data"];
        string? totalRows = wsData.Dimension.Address;
        ExcelRange data = wsData.Cells[totalRows];

        ExcelWorksheet? wsAuditPivot = p.Workbook.Worksheets.Add("Pivot2");

        ExcelPivotTable? pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit2");
        pivotTable1.ColumnGrandTotals = true;
        ExcelPivotTableField? rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);

        rowField.AddDateGrouping(eDateGroupBy.Years);
        ExcelPivotTableField? yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
        yearField.Name = "Year";

        _ = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

        ExcelPivotTableDataField? TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
        TotalSpend.Name = "Total Spend";
        TotalSpend.Format = "$##,##0";

        ExcelPivotTableDataField? CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
        CountInvoicePrice.Name = "Total Lines";
        CountInvoicePrice.Format = "##,##0";

        pivotTable1.DataOnRows = false;
    }

    [TestMethod]
    public void Issue15377()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("ws1");
        ws.Cells["A1"].Value = (double?)1;
        _ = ws.GetValue<double?>(1, 1);
    }

    [TestMethod]
    public void Issue15374()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("RT");
        ExcelRange? r = ws.Cells["A1"];
        r.RichText.Text = "Cell 1";
        _ = r["A2"].RichText.Add("Cell 2");
        SaveWorkbook(@"rt.xlsx", p);
    }

    [TestMethod]
    public void IssueTranslate()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Trans");
        ws.Cells["A1:A2"].Formula = "IF(1=1, \"A's B C\",\"D\") ";
        string? fr = ws.Cells["A1:A2"].FormulaR1C1;
        ws.Cells["A1:A2"].FormulaR1C1 = fr;
        Assert.AreEqual("IF(1=1,\"A's B C\",\"D\")", ws.Cells["A2"].Formula);
    }

    [TestMethod]
    public void Issue15397()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? workSheet = p.Workbook.Worksheets.Add("styleerror");
        workSheet.Cells["F:G"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["F:G"].Style.Fill.BackgroundColor.SetColor(Color.Red);

        workSheet.Cells["A:A,C:C"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["A:A,C:C"].Style.Fill.BackgroundColor.SetColor(Color.Red);

        //And then: 

        workSheet.Cells["A:H"].Style.Font.Color.SetColor(Color.Blue);

        workSheet.Cells["I:I"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["I:I"].Style.Fill.BackgroundColor.SetColor(Color.Red);
        workSheet.Cells["I2"].Style.Fill.BackgroundColor.SetColor(Color.Green);
        workSheet.Cells["I4"].Style.Fill.BackgroundColor.SetColor(Color.Blue);
        workSheet.Cells["I9"].Style.Fill.BackgroundColor.SetColor(Color.Pink);

        workSheet.InsertColumn(2, 2, 9);
        workSheet.Column(45).Width = 0;

        SaveWorkbook(@"styleerror.xlsx", p);
    }

    [TestMethod]
    public void Issuer14801()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? workSheet = p.Workbook.Worksheets.Add("rterror");
        ExcelRange? cell = workSheet.Cells["A1"];
        _ = cell.RichText.Add("toto: ");
        cell.RichText[0].PreserveSpace = true;
        cell.RichText[0].Bold = true;
        _ = cell.RichText.Add("tata");
        cell.RichText[1].Bold = false;
        cell.RichText[1].Color = Color.Green;
        SaveWorkbook(@"rtpreserve.xlsx", p);
    }

    [TestMethod]
    public void Issuer15445()
    {
        using ExcelPackage? p = new ExcelPackage();
        _ = p.Workbook.Worksheets.Add("ws1");
        ExcelWorksheet? ws2 = p.Workbook.Worksheets.Add("ws2");
        ws2.View.SelectedRange = "A1:B3 D12:D15";
        ws2.View.ActiveCell = "D15";
        SaveWorkbook(@"activeCell.xlsx", p);
    }

    [TestMethod]
    public void Issue15438()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Test");
        ExcelColor? c = ws.Cells["A1"].Style.Font.Color;
        c.Indexed = 3;
        Assert.AreEqual(c.LookupColor(c), "#FF00FF00");
    }

    public static byte[] ReadTemplateFile(string templateName)
    {
        byte[] templateFIle;

        using (MemoryStream ms = new MemoryStream())
        {
            using (FileStream? sw = new FileStream(templateName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                byte[] buffer = new byte[2048];
                int bytesRead;

                while ((bytesRead = sw.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, bytesRead);
                }
            }

            ms.Position = 0;
            templateFIle = ms.ToArray();
        }

        return templateFIle;
    }

    [TestMethod]
    public void Issue15455()
    {
        using ExcelPackage? pck = new ExcelPackage();

        ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("sheet1");
        ExcelWorksheet? sheet2 = pck.Workbook.Worksheets.Add("Sheet2");
        sheet1.Cells["C2"].Value = 3;
        sheet1.Cells["C3"].Formula = "VLOOKUP(E1, Sheet2!A1:D6, C2, 0)";
        sheet1.Cells["E1"].Value = "d";

        sheet2.Cells["A1"].Value = "d";
        sheet2.Cells["C1"].Value = "dg";
        pck.Workbook.Calculate();
        object? c3 = sheet1.Cells["C3"].Value;
        Assert.AreEqual("dg", c3);
    }

    [TestMethod]
    public void Issue15548_SumIfsShouldHandleGaps()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? test = package.Workbook.Worksheets.Add("Test");

        test.Cells["A1"].Value = 1;
        test.Cells["B1"].Value = "A";

        //test.Cells["A2"] is default
        test.Cells["B2"].Value = "A";

        test.Cells["A3"].Value = 1;
        test.Cells["B4"].Value = "B";

        test.Cells["D2"].Formula = "SUMIFS(A1:A3,B1:B3,\"A\")";

        test.Calculate();

        int result = test.Cells["D2"].GetValue<int>();

        Assert.AreEqual(1, result, string.Format("Expected 1, got {0}", result));
    }

    [TestMethod]
    public void Issue15548_SumIfsShouldHandleBadData()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? test = package.Workbook.Worksheets.Add("Test");

        test.Cells["A1"].Value = 1;
        test.Cells["B1"].Value = "A";

        test.Cells["A2"].Value = "Not a number";
        test.Cells["B2"].Value = "A";

        test.Cells["A3"].Value = 1;
        test.Cells["B4"].Value = "B";

        test.Cells["D2"].Formula = "SUMIFS(A1:A3,B1:B3,\"A\")";

        test.Calculate();

        int result = test.Cells["D2"].GetValue<int>();

        Assert.AreEqual(1, result, string.Format("Expected 1, got {0}", result));
    }

    [TestMethod]
    public void Issue63() // See https://github.com/JanKallman/EPPlus/issues/63
    {
        using ExcelPackage? p1 = new ExcelPackage();
        ExcelWorksheet ws = p1.Workbook.Worksheets.Add("ArrayTest");
        ws.Cells["A1"].Value = 1;
        ws.Cells["A2"].Value = 2;
        ws.Cells["A3"].Value = 3;
        ws.Cells["B1:B3"].CreateArrayFormula("A1:A3");
        p1.Save();

        // Test: basic support to recognize array formulas after reading Excel workbook file
        using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
        Assert.AreEqual("A1:A3", p1.Workbook.Worksheets["ArrayTest"].Cells["B1"].Formula);
        Assert.IsTrue(p1.Workbook.Worksheets["ArrayTest"].Cells["B1"].IsArrayFormula);
    }

    [TestMethod]
    public void Issue61()
    {
        DataTable table1 = new DataTable("TestTable");
        _ = table1.Columns.Add("name");
        _ = table1.Columns.Add("id");
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("i61");
        _ = ws.Cells["A1"].LoadFromDataTable(table1, true);
    }

    [TestMethod]
    public void Issue57()
    {
        ExcelPackage pck = new ExcelPackage();
        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("test");
        _ = ws.Cells["A1"].LoadFromArrays(Enumerable.Empty<object[]>());
    }

    [TestMethod]
    public void Issue66()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("Test!");
        ws.Cells["A1"].Value = 1;
        ws.Cells["B1"].Formula = "A1";
        ExcelWorkbook? wb = pck.Workbook;
        _ = wb.Names.Add("Name1", ws.Cells["A1:A2"]);
        _ = ws.Names.Add("Name2", ws.Cells["A1"]);
        pck.Save();
        using ExcelPackage? pck2 = new ExcelPackage(pck.Stream);
        _ = pck2.Workbook.Worksheets["Test!"];
    }

    /// <summary>
    /// Creating a new ExcelPackage with an external stream should not dispose of 
    /// that external stream. That is the responsibility of the caller.
    /// Note: This test would pass with EPPlus 4.1.1. In 4.5.1 the line CloseStream() was added
    /// to the ExcelPackage.Dispose() method. That line is redundant with the line before, 
    /// _stream.Close() except that _stream.Close() is only called if the _stream is NOT
    /// an External Stream (and several other conditions).
    /// Note that CloseStream() doesn't do anything different than _stream.Close().
    /// </summary>
    [TestMethod]
    public void Issue184_Disposing_External_Stream()
    {
        // Arrange
        MemoryStream? stream = new MemoryStream();

        using (ExcelPackage? excelPackage = new ExcelPackage(stream))
        {
            ExcelWorksheet? worksheet = excelPackage.Workbook.Worksheets.Add("Issue 184");
            worksheet.Cells[1, 1].Value = "Hello EPPlus!";
            excelPackage.SaveAs(stream);

            // Act
        } // This dispose should not dispose of stream.

        // Assert
        Assert.IsTrue(stream.Length > 0);
    }

    [TestMethod]
    public void Issue204()
    {
        using ExcelPackage? pack = new ExcelPackage();

        //create sheets
        ExcelWorksheet? sheet1 = pack.Workbook.Worksheets.Add("Sheet 1");
        ExcelWorksheet? sheet2 = pack.Workbook.Worksheets.Add("Sheet 2");

        //set some default values
        sheet1.Cells[1, 1].Value = 1;
        sheet2.Cells[1, 1].Value = 2;

        //fill the formula
        string? formula = string.Format("'{0}'!R1C1", sheet1.Name);

        ExcelRange? cell = sheet2.Cells[2, 1];
        cell.FormulaR1C1 = formula;

        //Formula should remain the same
        Assert.AreEqual(formula.ToUpper(), cell.FormulaR1C1.ToUpper());
    }

    //[TestMethod, Ignore]
    //public void Issue170()
    //{
    //    using ExcelPackage? p = OpenTemplatePackage("print_titles_170.xlsx");
    //    p.Compatibility.IsWorksheets1Based = false;
    //    ExcelWorksheet sheet = p.Workbook.Worksheets[0];

    //    sheet.PrinterSettings.RepeatColumns = new ExcelAddress("$A:$C");
    //    sheet.PrinterSettings.RepeatRows = new ExcelAddress("$1:$3");

    //    SaveWorkbook("print_titles_170-Saved.xlsx", p);
    //}

    [TestMethod]
    public void Issue172()
    {
        ExcelPackage? pck = OpenTemplatePackage("quest.xlsx");

        foreach (ExcelWorksheet? ws in pck.Workbook.Worksheets)
        {
            Console.WriteLine(ws.Name);
        }

        pck.Dispose();
    }

    [TestMethod]
    public void Issue219()
    {
        using ExcelPackage? p = OpenTemplatePackage("issueFile.xlsx");

        foreach (ExcelWorksheet? ws in p.Workbook.Worksheets)
        {
            Console.WriteLine(ws.Name);
        }
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidDataException))]
    public void Issue234()
    {
        using MemoryStream? s = new MemoryStream();
        byte[]? data = Encoding.UTF8.GetBytes("Bad data").ToArray();
        s.Write(data, 0, data.Length);
        _ = new ExcelPackage(s);
    }

    [TestMethod]
    public void WorksheetNameWithSingeQuote()
    {
        ExcelPackage? pck = OpenPackage("sheetname_pbl.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("Deal's History");
        _ = ws.Cells["A:B"];
        ws.AutoFilterAddress = ws.Cells["A1:C3"];
        _ = pck.Workbook.Names.Add("Test", ws.Cells["B1:D2"]);

        ExcelAddress? a2 = new ExcelAddress("'Deal''s History'!a1:a3");
        Assert.AreEqual(a2.WorkSheetName, "Deal's History");
        pck.Save();
        pck.Dispose();
    }

    [ExpectedException(typeof(ArgumentException))]
    [TestMethod]
    public void Issue233()
    {
        //get some test data
        List<Car>? cars = Car.GenerateList();

        ExcelPackage? pck = OpenPackage("issue233.xlsx", true);

        string? sheetName = "Summary_GLEDHOWSUGARCO![]()PTY";

        //Create the worksheet 
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add(sheetName);

        //Read the data into a range
        ExcelRangeBase? range = sheet.Cells["A1"].LoadFromCollection(cars, true);

        //Make the range a table
        ExcelTable? tbl = sheet.Tables.Add(range, $"data{sheetName}");
        tbl.ShowTotal = true;
        tbl.Columns["ReleaseYear"].TotalsRowFunction = RowFunctions.Sum;

        //save and dispose
        pck.Save();
        pck.Dispose();
    }

    public class Car
    {
        public int Id { get; set; }

        public string Make { get; set; }

        public string Model { get; set; }

        public int ReleaseYear { get; set; }

        public Car(int id, string make, string model, int releaseYear)
        {
            this.Id = id;
            this.Make = make;
            this.Model = model;
            this.ReleaseYear = releaseYear;
        }

        internal static List<Car> GenerateList() =>
            new()
            {
                //random data
                new Car(1, "Toyota", "Carolla", 1950),
                new Car(2, "Toyota", "Yaris", 2000),
                new Car(3, "Toyota", "Hilux", 1990),
                new Car(4, "Nissan", "Juke", 2010),
                new Car(5, "Nissan", "Trail Blazer", 1995),
                new Car(6, "Nissan", "Micra", 2018),
                new Car(7, "BMW", "M3", 1980),
                new Car(8, "BMW", "X5", 2008),
                new Car(9, "BMW", "M6", 2003),
                new Car(10, "Merc", "S Class", 2001)
            };
    }

    [TestMethod]
    public void Issue236()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue236.xlsx");
        _ = p.Workbook.Worksheets["Sheet1"].Cells[7, 10].AddComment("test", "Author");
        SaveWorkbook("Issue236-Saved.xlsx", p);
    }

    [TestMethod]
    public void Issue228()
    {
        using ExcelPackage? p = OpenTemplatePackage("Font55.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["Sheet1"];
        _ = ws.Drawings.AddShape("Shape1", eShapeStyle.Diamond);
        ws.Cells["A1"].Value = "tasetraser";
        ws.Cells.AutoFitColumns();
        SaveWorkbook("Font55-Saved.xlsx", p);
    }

    [TestMethod]
    public void Issue241()
    {
        ExcelPackage? pck = OpenPackage("issue241", true);
        ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("test");
        wks.DefaultRowHeight = 35;
        pck.Save();
        pck.Dispose();
    }

    [TestMethod]
    public void Issue195()
    {
        using ExcelPackage? pkg = new ExcelPackage();
        _ = pkg.Workbook.Worksheets.Add("Sheet1");
        ExcelNamedStyleXml? defaultStyle = pkg.Workbook.Styles.CreateNamedStyle("Default");
        defaultStyle.Style.Font.Name = "Arial";
        defaultStyle.Style.Font.Size = 18;
        defaultStyle.Style.Font.UnderLine = true;
        ExcelNamedStyleXml? boldStyle = pkg.Workbook.Styles.CreateNamedStyle("Bold", defaultStyle.Style);
        boldStyle.Style.Font.Color.SetColor(Color.Red);

        Assert.AreEqual("Arial", defaultStyle.Style.Font.Name);
        Assert.AreEqual(18, defaultStyle.Style.Font.Size);

        Assert.AreEqual("Arial", boldStyle.Style.Font.Name);
        Assert.AreEqual(18, boldStyle.Style.Font.Size);
        Assert.AreEqual(boldStyle.Style.Font.Color.Rgb, "FFFF0000");

        SaveWorkbook("DefaultStyle.xlsx", pkg);
    }

    [TestMethod]
    public void Issue332()
    {
        InitBase();
        ExcelPackage? pkg = OpenPackage("Hyperlink.xlsx", true);
        ExcelWorksheet? ws = pkg.Workbook.Worksheets.Add("Hyperlink");
        ws.Cells["A1"].Hyperlink = new ExcelHyperLink("A2", "A2");
        pkg.Save();
    }

    [TestMethod]
    public void Issue332_2()
    {
        InitBase();
        ExcelPackage? pkg = OpenPackage("Hyperlink.xlsx");
        ExcelWorksheet? ws = pkg.Workbook.Worksheets["Hyperlink"];
        Assert.IsNotNull(ws.Cells["A1"].Hyperlink);
    }

    [TestMethod]
    public void Issue347()
    {
        ExcelPackage? package = OpenTemplatePackage("Issue327.xlsx");
        ExcelWorksheet? templateWS = package.Workbook.Worksheets["Template"];

        //package.Workbook.Worksheets.Add("NewWs", templateWS);
        package.Workbook.Worksheets.Delete(templateWS);
    }

    [TestMethod]
    public void Issue348()
    {
        using ExcelPackage pck = new ExcelPackage();
        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("S1");
        string formula = "VLOOKUP(C2,A:B,1,0)";
        ws.Cells[2, 4].Formula = formula;
        ws.Cells[2, 5].FormulaR1C1 = ws.Cells[2, 4].FormulaR1C1;
    }

    [TestMethod]
    public void Issue367()
    {
        using ExcelPackage? pck = OpenTemplatePackage(@"ProductFunctionTest.xlsx");
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.First();

        //sheet.Cells["B13"].Value = null;
        sheet.Cells["B14"].Value = 11;
        sheet.Cells["B15"].Value = 13;
        sheet.Cells["B16"].Formula = "Product(B13:B15)";
        sheet.Calculate();

        Assert.AreEqual(0d, sheet.Cells["B16"].Value);
    }

    [TestMethod]
    public void Issue345()
    {
        using ExcelPackage package = OpenTemplatePackage("issue345.xlsx");
        ExcelWorksheet? worksheet = package.Workbook.Worksheets["test"];
        int[] sortColumns = new int[1];
        sortColumns[0] = 0;
        worksheet.Cells["A2:A30864"].Sort(sortColumns);
        package.Save();
    }

    [TestMethod]
    public void Issue387()
    {
        using ExcelPackage package = OpenTemplatePackage("issue345.xlsx");
        ExcelWorkbook? workbook = package.Workbook;
        ExcelWorksheet? worksheet = workbook.Worksheets.Add("One");

        worksheet.Cells[1, 3].Value = "Hello";
        ExcelRange? cells = worksheet.Cells["A3"];

        _ = worksheet.Names.Add("R0", cells);
        _ = workbook.Names.Add("Q0", cells);
    }

    [TestMethod]
    public void Issue333()
    {
        CultureInfo? ci = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-SE");

        using (ExcelPackage? package = new ExcelPackage())
        {
            ExcelWorksheet? ws = package.Workbook.Worksheets.Add("TextBug");
            ws.Cells["A1"].Value = new DateTime(2019, 3, 7);
            ws.Cells["A1"].Style.Numberformat.Format = "mm-dd-yy";

            Assert.AreEqual("2019-03-07", ws.Cells["A1"].Text);
        }

        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

        using (ExcelPackage? package = new ExcelPackage())
        {
            ExcelWorksheet? ws = package.Workbook.Worksheets.Add("TextBug");
            ws.Cells["A1"].Value = new DateTime(2019, 3, 7);
            ws.Cells["A1"].Style.Numberformat.Format = "mm-dd-yy";

            Assert.AreEqual("3/7/2019", ws.Cells["A1"].Text);
        }

        Thread.CurrentThread.CurrentCulture = ci;
    }

    [TestMethod]
    public void Issue445()
    {
        ExcelPackage p = new ExcelPackage();
        ExcelWorksheet ws = p.Workbook.Worksheets.Add("AutoFit"); //<-- This line takes forever. The process hangs.
        ws.Cells[1, 1].Value = new string('a', 50000);
        ws.Cells[1, 1].AutoFitColumns();
    }

    [TestMethod]
    public void Issue551()
    {
        using ExcelPackage? p = OpenTemplatePackage("Submittal.Extract.5.ton.xlsx");
        SaveWorkbook("Submittal.Extract.5.ton_Saved.xlsx", p);
    }

    [TestMethod]
    public void Issue558()
    {
        using ExcelPackage? p = OpenTemplatePackage("GoogleSpreadsheet.xlsx");
        ExcelWorksheet ws = p.Workbook.Worksheets[0];
        _ = p.Workbook.Worksheets.Copy(ws.Name, "NewName");
        SaveWorkbook("GoogleSpreadsheet-Saved.xlsx", p);
    }

    [TestMethod]
    public void Issue520()
    {
        using ExcelPackage? p = OpenTemplatePackage("template_slim.xlsx");

        ExcelWorksheet? workSheet = p.Workbook.Worksheets[0];
        _ = workSheet.Cells["B5"].LoadFromArrays(new List<object[]> { new object[] { "xx", "Name", 1, 2, 3, 5, 6, 7 } });

        SaveWorkbook("ErrorStyle0.xlsx", p);
    }

    [TestMethod]
    public void Issue510()
    {
        using ExcelPackage? p = OpenTemplatePackage("Error.Opening.with.EPPLus.xlsx");

        SaveWorkbook("Issue510.xlsx", p);
    }

    [TestMethod]
    public void Issue464()
    {
        using ExcelPackage? p1 = OpenTemplatePackage("Sample_Cond_Format.xlsx");
        ExcelWorksheet? ws = p1.Workbook.Worksheets[0];
        using ExcelPackage? p2 = new ExcelPackage();
        ExcelWorksheet? ws2 = p2.Workbook.Worksheets.Add("Test", ws);

        foreach (IExcelConditionalFormattingRule? cf in ws2.ConditionalFormatting)
        {
        }

        SaveWorkbook("CondCopy.xlsx", p2);
    }

    [TestMethod]
    public void Issue436()
    {
        using ExcelPackage? p1 = OpenTemplatePackage("issue436.xlsx");
        ExcelWorksheet? ws = p1.Workbook.Worksheets[0];
        Assert.IsNotNull(((ExcelShape)ws.Drawings[0]).Text);
    }

    [TestMethod]
    public void Issue425()
    {
        using ExcelPackage? p1 = OpenTemplatePackage("issue425.xlsm");
        ExcelWorksheet? ws = p1.Workbook.Worksheets[1];

        _ = p1.Workbook.Worksheets.Add("NewNotCopied");
        _ = p1.Workbook.Worksheets.Add("NewCopied", ws);

        SaveWorkbook("issue425.xlsm", p1);
    }

    [TestMethod]
    public void Issue422()
    {
        using ExcelPackage? p1 = OpenTemplatePackage("CustomFormula.xlsx");
        SaveWorkbook("issue422.xlsx", p1);
    }

    [TestMethod]
    public void Issue625()
    {
        using ExcelPackage? p = OpenTemplatePackage("multiple_print_areas.xlsx");

        SaveWorkbook("Issue625.xlsx", p);
    }

    [TestMethod]
    public void Issue403()
    {
        using ExcelPackage? p = OpenTemplatePackage("issue403.xlsx");
        SaveWorkbook("Issue403.xlsx", p);
    }

    [TestMethod]
    public void Issue39()
    {
        using ExcelPackage? p = OpenTemplatePackage("MyExcel.xlsx");
        ExcelWorksheet? workSheet = p.Workbook.Worksheets[0];

        workSheet.InsertRow(8, 2, 8);
        SaveWorkbook("Issue39.xlsx", p);
    }

    [TestMethod]
    public void Issue70()
    {
        using ExcelPackage? p = OpenTemplatePackage("HiddenOO.xlsx");
        Assert.IsTrue(p.Workbook.Worksheets[0].Column(2).Hidden);
        SaveWorkbook("Issue70.xlsx", p);
    }

    [TestMethod]
    public void Issue72()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue72-Table.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        Assert.AreEqual("COUNTIF(Base[Date],Calc[[#This Row],[Date]])", ws.Cells["F3"].Formula);
        SaveWorkbook("Issue72.xlsx", p);
    }

    [TestMethod]
    public void Issue54()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("MergeBug");

        ExcelRange? r = ws.Cells[1, 1, 1, 5];
        r.Merge = true;
        r.Value = "Header";
        SaveWorkbook("Issue54.xlsx", p);
    }

    [TestMethod]
    public void Issue55()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? worksheet = p.Workbook.Worksheets.Add("DV test");
        string? rangeToSet = ExcelCellBase.GetAddress(1, 3, ExcelPackage.MaxRows, 3);
        _ = worksheet.Names.Add("ListName", worksheet.Cells["D1:D3"]);
        worksheet.Cells["D1"].Value = "A";
        worksheet.Cells["D2"].Value = "B";
        worksheet.Cells["D3"].Value = "C";
        IExcelDataValidationList? validation = worksheet.DataValidations.AddListValidation(rangeToSet);
        validation.Formula.ExcelFormula = $"=ListName";
        SaveWorkbook("dv.xlsx", p);
    }

    [TestMethod]
    public void Issue73()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue73.xlsx");

        SaveWorkbook("Issue73Saved.xlsx", p);
    }

    [TestMethod]
    public void Issue74()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue74.xlsx");

        SaveWorkbook("Issue74Saved.xlsx", p);
    }

    [TestMethod]
    public void Issue76()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue76.xlsx");

        SaveWorkbook("Issue76Saved.xlsx", p);
    }

    [TestMethod]
    public void Issue88()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue88.xlsm");
        ExcelWorksheet? ws1 = p.Workbook.Worksheets[0];
        _ = p.Workbook.Worksheets.Add("test", ws1);
        SaveWorkbook("Issue88Saved.xlsm", p);
    }

    [TestMethod]
    public void Issue94()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue425.xlsm");
        p.Workbook.VbaProject.Remove();
        SaveWorkbook("Issue425.xlsx", p);
    }

    [TestMethod]
    public void Issue95()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue95.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue99()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue-99-2.xlsx");

        //var p2 = OpenPackage("Issue99-2Saved-new.xlsx", true);
        //var ws = p2.Workbook.Worksheets.Add("Picture");
        //ws.Drawings.AddPicture("Test1", Properties.Resources.Test1);
        //p.Workbook.Worksheets.Add("copy1", p.Workbook.Worksheets[0]);
        //p2.Workbook.Worksheets.Add("copy1", p.Workbook.Worksheets[0]);
        //p.Workbook.Worksheets.Add("copy2", p2.Workbook.Worksheets[0]);
        //SaveAndCleanup(p2);
        SaveWorkbook("Issue99-2Saved.xlsx", p);
    }

    [TestMethod]
    public void Issue115()
    {
        using ExcelPackage? p = OpenPackage("Issue115.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("DefinedNamesIssue");
        _ = p.Workbook.Names.Add("Name", ws.Cells["B6:D8,B10:D11"]);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue121()
    {
        using ExcelPackage? p = OpenTemplatePackage("Deployment aftaler.xlsx");
    }

    //[TestMethod, Ignore]
    //public void SupportCase17()
    //{
    //    using ExcelPackage? p = new ExcelPackage(new FileInfo(@"c:\temp\Issue17\BreakLinks3.xlsx"));
    //    Stopwatch? stopwatch = Stopwatch.StartNew();
    //    p.Workbook.FormulaParserManager.AttachLogger(new FileInfo("c:\\temp\\formulalog.txt"));
    //    p.Workbook.Calculate();
    //    stopwatch.Stop();
    //    double ms = stopwatch.Elapsed.TotalSeconds;
    //}

    //[TestMethod, Ignore]
    //public void Issue17()
    //{
    //    using ExcelPackage? p = OpenTemplatePackage("Excel Sample Circular Ref break links.xlsx");
    //    p.Workbook.Calculate();
    //}

    [TestMethod]
    public void Issue18()
    {
        using ExcelPackage? p = OpenTemplatePackage("000P020-SQ101_H0.xlsm");
        p.Workbook.Worksheets.Delete(0);
        p.Workbook.Worksheets.Delete(2);
        p.Workbook.Calculate();
        SaveWorkbook("null_issue_vba.xlsm", p);
    }

    [TestMethod]
    public void Issue26()
    {
        using (ExcelPackage? p = OpenTemplatePackage("Issue26.xlsx"))
        {
            SaveAndCleanup(p);
        }

        using (ExcelPackage? p = OpenPackage("Issue26.xlsx"))
        {
            SaveWorkbook("Issue26-resaved.xlsx", p);
        }
    }

    [TestMethod]
    public void Issue180()
    {
        ExcelPackage? p1 = OpenTemplatePackage("Issue180-1.xlsm");
        ExcelPackage? p2 = OpenTemplatePackage("Issue180-2.xlsm");
        _ = p2.Workbook.Worksheets.Add(p1.Workbook.Worksheets[0].Name, p1.Workbook.Worksheets[0]);
        p2.SaveAs(new FileInfo("c:\\epplustest\\t.xlsm"));
    }

    [TestMethod]
    public void Issue34()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue34.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue38()
    {
        using ExcelPackage? p = OpenTemplatePackage("pivottest.xlsx");
        Assert.AreEqual(1, p.Workbook.Worksheets[1].PivotTables.Count);
        ExcelTable? tbl = p.Workbook.Worksheets[0].Tables[0];
        ExcelPivotTable? pt = p.Workbook.Worksheets[1].PivotTables[0];
        Assert.IsNotNull(p.Workbook.Worksheets[1].PivotTables[0].CacheDefinition);
        ExcelPivotTableSlicer? s1 = pt.Fields[0].AddSlicer();
        s1.SetPosition(0, 500);
        ExcelPivotTableSlicer? s2 = pt.Fields["OpenDate"].AddSlicer();
        pt.Fields["Distance"].Format = "#,##0.00";
        _ = pt.Fields["Distance"].AddSlicer();
        s2.SetPosition(0, 500 + (int)s1._width);
        _ = tbl.Columns["IsUser"].AddSlicer();
        _ = pt.Fields["IsUser"].AddSlicer();

        SaveWorkbook("pivotTable2.xlsx", p);
    }

    [TestMethod]
    public void Issue195_PivotTable()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue195.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue45()
    {
        using ExcelPackage? p = OpenPackage("LinkIssue.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1:A2"].Value = 1;
        ws.Cells["B1:B2"].Formula = $"VLOOKUP($A1,[externalBook.xlsx]Prices!$A:$H, 3, FALSE)";
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void EmfIssue()
    {
        using ExcelPackage? p = OpenTemplatePackage("emfIssue.xlsm");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue201()
    {
        using ExcelPackage? p = OpenTemplatePackage("book1.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        Assert.AreEqual("0", ws.Cells["A1"].Text);
        Assert.AreEqual("-", ws.Cells["A2"].Text);
        Assert.AreEqual("0", ws.Cells["A3"].Text);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void IssueCellstore()
    {
        int START_ROW = 1;
        int CustomTemplateRowsOffset = 4;
        int rowCount = 34000;
        using ExcelPackage? package = OpenTemplatePackage("CellStoreIssue.xlsm");
        ExcelWorksheet? worksheet = package.Workbook.Worksheets[0];
        worksheet.Cells["A5"].Value = "Test";
        worksheet.InsertRow(START_ROW + CustomTemplateRowsOffset, rowCount - 1, CustomTemplateRowsOffset + 1);
        Assert.AreEqual("Test", worksheet.Cells["A34004"].Value);

        //for (int k = START_ROW+CustomTemplateRowsOffset; k < rowCount; k++)
        //{
        //    worksheet.Cells[(START_ROW + CustomTemplateRowsOffset) + ":" + (START_ROW + CustomTemplateRowsOffset)]
        //        .Copy(worksheet.Cells[k + 1 + ":" + k + 1]);
        //}
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void Issue220()
    {
        using ExcelPackage? p = OpenTemplatePackage("Generated.with.EPPlus.xlsx");
    }

    [TestMethod]
    public void Issue232()
    {
        using ExcelPackage? p = OpenTemplatePackage("pivotbug541.xlsx");
        ExcelWorksheet? overviewSheet = p.Workbook.Worksheets["Overblik"];
        ExcelWorksheet? serverSheet = p.Workbook.Worksheets["Servers"];

        _ = overviewSheet.PivotTables.Add(overviewSheet.Cells["A4"], serverSheet.Cells[serverSheet.Dimension.Address], "ServerPivot");

        p.Save();
    }

    [TestMethod]
    public void Issue_234()
    {
        using ExcelPackage? p = OpenTemplatePackage("ExcelErrorFile.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["Leistung"];

        Assert.IsNull(ws.Cells["C65538"].Value);
        Assert.IsNull(ws.Cells["C71715"].Value);
        Assert.AreEqual(0D, ws.Cells["C71716"].Value);
        Assert.AreEqual(0D, ws.Cells["C71811"].Value);
        Assert.IsNull(ws.Cells["C71812"].Value);
        Assert.IsNull(ws.Cells["C77667"].Value);
        Assert.AreEqual(0D, ws.Cells["C77668"].Value);
    }

    [TestMethod]
    public void InflateIssue()
    {
        using ExcelPackage? p = OpenPackage("inflateStart.xlsx", true);
        ExcelWorksheet? worksheet = p.Workbook.Worksheets.Add("Test");

        for (int i = 1; i <= 10; i++)
        {
            worksheet.Cells[1, i].Hyperlink = new Uri("https://epplussoftware.com");
            worksheet.Cells[1, i].Value = "Url " + worksheet.Cells[1, i].Address;
        }

        p.Save();
        using ExcelPackage? p2 = new ExcelPackage(p.Stream);

        for (int i = 0; i < 10; i++)
        {
            p.Save();
        }

        SaveWorkbook("Inflate.xlsx", p2);
    }

    [TestMethod]
    public void DrawingSetFont()
    {
        using ExcelPackage? p = OpenPackage("DrawingSetFromFont.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Drawing1");
        ExcelShape? shape = ws.Drawings.AddShape("x", eShapeStyle.Rect);
        shape.Font.SetFromFont("Arial", 20);
        shape.Text = "Font";
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue_258()
    {
        using ExcelPackage? package = OpenTemplatePackage("Test.xlsx");
        ExcelWorksheet? overviewSheet = package.Workbook.Worksheets["Overview"];

        if (overviewSheet != null)
        {
            package.Workbook.Worksheets.Delete(overviewSheet);
        }

        overviewSheet = package.Workbook.Worksheets.Add("Overview");
        ExcelWorksheet? serverSheet = package.Workbook.Worksheets["Servers"];

        ExcelPivotTable? serverPivot =
            overviewSheet.PivotTables.Add(overviewSheet.Cells["A4"], serverSheet.Cells[serverSheet.Dimension.Address], "ServerPivot");

        ExcelPivotTableField? serverNameField = serverPivot.Fields["Name"];
        _ = serverPivot.RowFields.Add(serverNameField);
        ExcelPivotTableField? standardBackupField = serverPivot.Fields["StandardBackup"];
        _ = serverPivot.PageFields.Add(standardBackupField);
        standardBackupField.Items.Refresh();
        ExcelPivotTableFieldItemsCollection? items = standardBackupField.Items;
        items.SelectSingleItem(1); // <===== this one is to select only the "false" condition
        SaveWorkbook("Issue248.xlsx", package);
    }

    [TestMethod]
    public void Issue_243()
    {
        using ExcelPackage? p = OpenPackage("formula.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("formula");
        ws.Cells["A1"].Value = "column1";
        ws.Cells["A2"].Value = 1;
        ws.Cells["A3"].Value = 2;
        ws.Cells["A4"].Value = 3;

        _ = ws.Tables.Add(ws.Cells["A1:A4"], "Table1");

        ws.Cells["B1"].Formula = "TEXTJOIN(\" | \", false, INDIRECT(\"Table1[#data]\"))";
        ws.Calculate();
        Assert.AreEqual("1 | 2 | 3", ws.Cells["B1"].Value);

        ws.Cells["B1"].Formula = "TEXTJOIN(\" | \", false, INDIRECT(\"Table1\"))";
        ws.Calculate();
        Assert.AreEqual("1 | 2 | 3", ws.Cells["B1"].Value);
    }

    [TestMethod]
    public void IssueCommentInsert()
    {
        using ExcelPackage? p = OpenPackage("CommentInsert.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("CommentInsert");
        _ = ws.Cells["A2"].AddComment("na", "test");
        Assert.AreEqual(1, ws.Comments.Count);

        ws.InsertRow(2, 1);
        ws.Cells["A3"].Insert(eShiftTypeInsert.Right);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue261()
    {
        using ExcelPackage? p = OpenTemplatePackage("issue261.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["data"];
        ws.Cells["A1"].Value = "test";
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue260()
    {
        using ExcelPackage? p = OpenTemplatePackage("issue260.xlsx");
        ExcelWorkbook? workbook = p.Workbook;
        Console.WriteLine(workbook.Worksheets.Count);
    }

    [TestMethod]
    public void Issue268()
    {
        using ExcelPackage? p = OpenPackage("Issue268.xlsx", true);
        ExcelWorksheet formSheet = CreateFormSheet(p);
        ExcelControlCheckBox? r1 = formSheet.Drawings.AddCheckBoxControl("OptionSingleRoom");
        r1.Text = "Single Room";
        r1.LinkedCell = formSheet.Cells["G7"];
        r1.SetPosition(5, 0, 1, 0);
        _ = p.Workbook.Worksheets.Add("Table");
        ExcelRange tableRange = formSheet.Cells[10, 20, 30, 22];
        ExcelTable faultsTable = formSheet.Tables.Add(tableRange, "FaultsTable");
        faultsTable.StyleName = "None";
        SaveAndCleanup(p);
    }

    private static ExcelWorksheet CreateFormSheet(ExcelPackage package)
    {
        ExcelWorksheet? formSheet = package.Workbook.Worksheets.Add("Form");
        formSheet.Cells["A1"].Value = "Room booking";
        formSheet.Cells["A1"].Style.Font.Size = 18;
        formSheet.Cells["A1"].Style.Font.Bold = true;

        return formSheet;
    }

    [TestMethod]
    public void Issue269()
    {
        List<TestDTO>? data = new List<TestDTO>();

        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? sheet = p.Workbook.Worksheets.Add("Sheet1");
        ExcelRangeBase? r = sheet.Cells["A1"].LoadFromCollection(data, false);
        Assert.IsNull(r);
    }

    [TestMethod]
    public void Issue272()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue272.xlsx");
        ExcelWorkbook? workbook = p.Workbook;
        Console.WriteLine(workbook.Worksheets.Count);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void IssueS84()
    {
        using ExcelPackage? p = OpenTemplatePackage("XML in Cells.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ExcelRange? cell = ws.Cells["D43"];
        cell.Value += " ";

        ExcelRichText rtx = cell.RichText.Add("a");

        rtx.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;

        ws.Cells["D43:E44"].Value = new object[,] { { "Cell1", "Cell2" }, { "Cell21", "Cell22" } };

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void IssueS80()
    {
        using ExcelPackage? p = OpenTemplatePackage("Example - CANNOT OPEN EPPLUS.xlsx");
        SaveAndCleanup(p);
        //_ = new ExcelAddress("f");
    }

    [TestMethod]
    public void IssueS91()
    {
        using ExcelPackage? p = OpenTemplatePackage("Tagging Template V14.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["Stacked Logs"];

        //Insert 2 rows extending the data validations. 
        ws.InsertRow(4, 2, 4);

        //Get the data validation of choice.
        IExcelDataValidationList? dv = ws.DataValidations[0].As.ListValidation;

        //Adjust the formula using the R1C1 translator...
        string? formula = dv.Formula.ExcelFormula;
        string? r1c1Formula = OfficeOpenXml.Core.R1C1Translator.ToR1C1Formula(formula, dv.Address.Start.Row, dv.Address.Start.Column);

        //Add one row to the formula
        _ = OfficeOpenXml.Core.R1C1Translator.FromR1C1Formula(r1c1Formula, dv.Address.Start.Row + 1, dv.Address.Start.Column);

        SaveAndCleanup(p);
    }

    public class Test
    {
        public int Value1 { get; set; }

        public int Value2 { get; set; }

        public int Value3 { get; set; }
    }

    [TestMethod]
    public void Issue284()
    {
        //1
        List<Test>? report1 = new List<Test>
        {
            new Test { Value1 = 1, Value2 = 2, Value3 = 3 },
            new Test { Value1 = 2, Value2 = 3, Value3 = 4 },
            new Test { Value1 = 5, Value2 = 6, Value3 = 7 }
        };

        //3
        List<Test>? report2 = new List<Test>
        {
            new Test { Value1 = 0, Value2 = 0, Value3 = 0 },
            new Test { Value1 = 0, Value2 = 0, Value3 = 0 },
            new Test { Value1 = 0, Value2 = 0, Value3 = 0 }
        };

        //4
        List<Test>? report3 = new List<Test>
        {
            new Test { Value1 = 3, Value2 = 3, Value3 = 3 },
            new Test { Value1 = 3, Value2 = 3, Value3 = 3 },
            new Test { Value1 = 3, Value2 = 3, Value3 = 3 }
        };

        using ExcelPackage? excelFile = OpenTemplatePackage("issue284.xlsx");

        //Data1
        ExcelWorksheet? worksheet = excelFile.Workbook.Worksheets["Test1"];
        ExcelRangeBase location = worksheet.Cells["A1"].LoadFromCollection(Collection: report1, PrintHeaders: true);
        ExcelTable? t = worksheet.Tables.Add(location, "mytestTbl");
        t.TableStyle = TableStyles.None;

        //Data2
        worksheet = excelFile.Workbook.Worksheets["Test2"];
        location = worksheet.Cells["A1"].LoadFromCollection(Collection: report2, PrintHeaders: true);
        _ = worksheet.Tables.Add(location, "mytestsureTbl");

        //Data3
        location = worksheet.Cells["K1"].LoadFromCollection(Collection: report3, PrintHeaders: true);
        _ = worksheet.Tables.Add(location, "Test3");

        ExcelWorksheet? wsFirst = excelFile.Workbook.Worksheets["Test1"];

        wsFirst.Select();
        SaveAndCleanup(excelFile);
    }

    [TestMethod]
    public void Ticket90()
    {
        using ExcelPackage? p = OpenTemplatePackage("Example - Calculate.xlsx");
        ExcelWorksheet? sheet = p.Workbook.Worksheets["Others"];
        FileInfo? fi = new FileInfo(@"c:\Temp\countiflog.txt");
        p.Workbook.FormulaParserManager.AttachLogger(fi);
        sheet.Calculate(x => x.PrecisionAndRoundingStrategy = OfficeOpenXml.FormulaParsing.PrecisionAndRoundingStrategy.Excel);
        p.Workbook.FormulaParserManager.DetachLogger();
        //_ = new ExcelAddress();

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Ticket90_2()
    {
        using ExcelPackage? p = OpenTemplatePackage("s70.xlsx");
        p.Workbook.Calculate();
        Assert.AreEqual(7D, p.Workbook.Worksheets[0].Cells["P1"].Value);
        Assert.AreEqual(1D, p.Workbook.Worksheets[0].Cells["P2"].Value);
        Assert.AreEqual(0D, p.Workbook.Worksheets[0].Cells["P3"].Value);
    }

    [TestMethod]
    public void Issue287()
    {
        using ExcelPackage? p = OpenTemplatePackage("issue287.xlsm");
        p.Workbook.CreateVBAProject();
        p.Save();
    }

    [TestMethod]
    public void Issue309()
    {
        using ExcelPackage? p = OpenTemplatePackage("test1.xlsx");
        p.Save();
    }

    [TestMethod]
    public void Issue274()
    {
        using ExcelPackage? p = OpenPackage("Issue274.xlsx", true);
        ExcelWorksheet? worksheet = p.Workbook.Worksheets.Add("PayrollData");

        //Add the headers
        worksheet.Cells[1, 1].Value = "Employee";
        worksheet.Cells[1, 2].Value = "HomeOffice";
        worksheet.Cells[1, 3].Value = "JobNo";
        worksheet.Cells[1, 4].Value = "Ordinary";
        worksheet.Cells[1, 5].Value = "TimeHalf";
        worksheet.Cells[1, 6].Value = "DoubleTime";
        worksheet.Cells[1, 7].Value = "ProductiveHrs";
        worksheet.Cells[1, 8].Value = "NonProductiveHrs";

        int cnt = 2;
        worksheet.Cells[cnt, 1].Value = "Steve";
        worksheet.Cells[cnt, 2].Value = "Binda";
        worksheet.Cells[cnt, 3].Value = "SW001";
        worksheet.Cells[cnt, 4].Value = 12.0;
        worksheet.Cells[cnt, 5].Value = 6.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 18.0;
        worksheet.Cells[cnt, 8].Value = 0.0;
        cnt++;
        worksheet.Cells[cnt, 1].Value = "Steve";
        worksheet.Cells[cnt, 2].Value = "Binda";
        worksheet.Cells[cnt, 3].Value = "SW002";
        worksheet.Cells[cnt, 4].Value = 7.0;
        worksheet.Cells[cnt, 5].Value = 0.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 7.0;
        worksheet.Cells[cnt, 8].Value = 0.0;
        cnt++;
        worksheet.Cells[cnt, 1].Value = "Steve";
        worksheet.Cells[cnt, 2].Value = "Binda";
        worksheet.Cells[cnt, 3].Value = "Admin";
        worksheet.Cells[cnt, 4].Value = 4.0;
        worksheet.Cells[cnt, 5].Value = 0.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 0.0;
        worksheet.Cells[cnt, 8].Value = 4.0;
        cnt++;
        worksheet.Cells[cnt, 1].Value = "Peter";
        worksheet.Cells[cnt, 2].Value = "Binda";
        worksheet.Cells[cnt, 3].Value = "SW001";
        worksheet.Cells[cnt, 4].Value = 12.0;
        worksheet.Cells[cnt, 5].Value = 6.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 18.0;
        worksheet.Cells[cnt, 8].Value = 0.0;
        cnt++;
        worksheet.Cells[cnt, 1].Value = "Peter";
        worksheet.Cells[cnt, 2].Value = "Binda";
        worksheet.Cells[cnt, 3].Value = "SW002";
        worksheet.Cells[cnt, 4].Value = 7.0;
        worksheet.Cells[cnt, 5].Value = 0.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 7.0;
        worksheet.Cells[cnt, 8].Value = 0.0;
        cnt++;
        worksheet.Cells[cnt, 1].Value = "Peter";
        worksheet.Cells[cnt, 2].Value = "Binda";
        worksheet.Cells[cnt, 3].Value = "Admin";
        worksheet.Cells[cnt, 4].Value = 4.0;
        worksheet.Cells[cnt, 5].Value = 0.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 0.0;
        worksheet.Cells[cnt, 8].Value = 4.0;
        cnt++;
        worksheet.Cells[cnt, 1].Value = "Brian";
        worksheet.Cells[cnt, 2].Value = "Sydney";
        worksheet.Cells[cnt, 3].Value = "SW001";
        worksheet.Cells[cnt, 4].Value = 12.0;
        worksheet.Cells[cnt, 5].Value = 6.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 18.0;
        worksheet.Cells[cnt, 8].Value = 0.0;
        cnt++;
        worksheet.Cells[cnt, 1].Value = "Brian";
        worksheet.Cells[cnt, 2].Value = "Binda";
        worksheet.Cells[cnt, 3].Value = "SW002";
        worksheet.Cells[cnt, 4].Value = 7.0;
        worksheet.Cells[cnt, 5].Value = 0.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 7.0;
        worksheet.Cells[cnt, 8].Value = 0.0;
        cnt++;
        worksheet.Cells[cnt, 1].Value = "Brian";
        worksheet.Cells[cnt, 2].Value = "Binda";
        worksheet.Cells[cnt, 3].Value = "Admin";
        worksheet.Cells[cnt, 4].Value = 4.0;
        worksheet.Cells[cnt, 5].Value = 0.0;
        worksheet.Cells[cnt, 6].Value = 0.0;
        worksheet.Cells[cnt, 7].Value = 0.0;
        worksheet.Cells[cnt, 8].Value = 4.0;

        cnt--;

        using (ExcelRange? range = worksheet.Cells[1, 1, 1, 8])
        {
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            range.Style.Font.Color.SetColor(Color.White);
        }

        ExcelRange? dataRange = worksheet.Cells[1, 1, cnt, 8];
        ExcelTableCollection tblcollection = worksheet.Tables;
        ExcelTable table = tblcollection.Add(dataRange, "payrolldata");
        table.ShowHeader = true;
        table.ShowFilter = true;
        ExcelWorksheet? wsPivot = p.Workbook.Worksheets.Add("Employee-Job");
        ExcelPivotTable? pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], dataRange, "ByEmployee");
        _ = pivotTable.RowFields.Add(pivotTable.Fields["Employee"]);
        ExcelPivotTableField? rowField1 = pivotTable.RowFields.Add(pivotTable.Fields["HomeOffice"]);
        _ = pivotTable.RowFields.Add(pivotTable.Fields["JobNo"]);
        ExcelPivotTableField? calcField1 = pivotTable.Fields.AddCalculatedField("Productive", "'ProductiveHrs'/('ProductiveHrs'+'NonProductiveHrs')*100");
        calcField1.Format = "#,##0";
        ExcelPivotTableDataField dataField = pivotTable.DataFields.Add(pivotTable.Fields["Productive"]);
        dataField.Format = "#,##0.0";
        dataField.Name = "Productive2";

        dataField = pivotTable.DataFields.Add(pivotTable.Fields["Ordinary"]);
        dataField.Format = "#,##0.0";
        dataField = pivotTable.DataFields.Add(pivotTable.Fields["TimeHalf"]);
        dataField.Format = "#,##0.0";
        dataField = pivotTable.DataFields.Add(pivotTable.Fields["DoubleTime"]);
        dataField.Format = "#,##0.0";
        dataField = pivotTable.DataFields.Add(pivotTable.Fields["ProductiveHrs"]);
        dataField.Format = "#,##0.0";
        dataField = pivotTable.DataFields.Add(pivotTable.Fields["NonProductiveHrs"]);
        dataField.Format = "#,##0.0";
        pivotTable.DataOnRows = false;
        pivotTable.Compact = true;
        pivotTable.CompactData = true;
        pivotTable.OutlineData = true;

        //pivotTable.ShowDrill = true;
        //pivotTable.CacheDefinition.Refresh();
        pivotTable.Fields["Employee"].Items.ShowDetails(false);
        rowField1.Items.ShowDetails(false);
        worksheet.Cells.AutoFitColumns(0);

        // create macro's to collapse pivot table

        //p.Workbook.CreateVBAProject();
        //var sb = new StringBuilder();
        //sb.AppendLine("Private Sub Workbook_Open()");
        //sb.AppendLine("    Sheets(\"Employee-Job\").Select");
        //sb.AppendLine("    ActiveSheet.PivotTables(\"ByEmployee\").PivotFields(\"Employee\").ShowDetail = False");
        //sb.AppendLine("    ActiveSheet.PivotTables(\"ByEmployee\").PivotFields(\"HomeOffice\").ShowDetail = False");
        //sb.AppendLine("End Sub");
        //p.Workbook.CodeModule.Code = sb.ToString();
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void DeleteCommentIssue()
    {
        using ExcelPackage? p = OpenTemplatePackage("CommentDelete.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["S3"];
        _ = ws.Comments[0].RichText.Add("T");
        _ = ws.Comments[0].RichText.Add("x");
        ExcelWorksheet? ws2 = p.Workbook.Worksheets.Add("Copied S3", ws);
        ws.InsertRow(2, 1);
        ws.DeleteRow(2);
        ws2.DeleteRow(2);
        ws.InsertRow(2, 2);
        ExcelCommentCollection? c = ws2.Comments; // Access the comment collection to force loading it. Otherwise Exception!
        int dummy = c.Count; // to load!
        p.Workbook.Worksheets.Delete(ws);
        p.Workbook.Worksheets.Delete(ws2);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void DeleteWorksheetIssue()
    {
        using ExcelPackage? p = OpenTemplatePackage("CommentDelete.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["S3"];
        ExcelCommentCollection? c = ws.Comments; // Access the comment collection to force loading it. Otherwise Exception!
        int dummy = c.Count; // to load!

        //ws.DeleteRow(2);
        //dummy = c.Count; // to load!
        p.Workbook.Worksheets.Delete(ws);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue294()
    {
        using ExcelPackage? p = OpenTemplatePackage("test_excel_workbook_before2-xl.xlsx");
        p.Save();
    }

    [TestMethod]
    public void Issue333_2()
    {
        using ExcelPackage? p = OpenTemplatePackage("issue333-2.xlsx");
        ExcelWorksheet? sheet = p.Workbook.Worksheets[1];
        Assert.IsFalse(string.IsNullOrEmpty(sheet.Cells[8, 1].Formula));
        Assert.IsFalse(string.IsNullOrEmpty(sheet.Cells[9, 1].Formula));
        Assert.IsFalse(string.IsNullOrEmpty(sheet.Cells[9, 2].Formula));
        Assert.IsFalse(string.IsNullOrEmpty(sheet.Cells[32, 2].Formula));
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void S107()
    {
        using ExcelPackage? p = OpenTemplatePackage("2021-03-18 - Styling issues.xlsm");
        p.Save();
        ExcelPackage? p2 = new ExcelPackage(p.Stream);
        p2.SaveAs(new FileInfo(p.File.DirectoryName + "\\Test.xlsm"));
    }

    [TestMethod]
    public void S127()
    {
        using ExcelPackage? p = OpenTemplatePackage("Tagging Template V15 - New Format.xlsx");
        SaveWorkbook("Tagging Template V15 - New Format2.xlsx", p);
    }

    [TestMethod]
    public void MergeIssue()
    {
        using ExcelPackage? p = OpenTemplatePackage("MergeIssue.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["s7"];
        ws.Cells["B2:F2"].Merge = false;
        ws.Cells["B2:F12"].Clear();
        ws.Cells["B2:F2"].Merge = true;

        ws.Cells["B2:F12"].Merge = false;
        ws.Cells["B2:F12"].Clear();

        ws.Cells["B2:F12"].Merge = true;
        ws.Cells["B1:F12"].Clear();

        ws.Cells["B2:F2"].Merge = true;
        ws.Cells["B2:F2"].Merge = false;
        ws.Cells["B2:F12"].Clear();
        ws.Cells["B2:F2"].Merge = true;
    }

    public static void DefinedNamesAddressIssue()
    {
        using ExcelPackage? p = OpenPackage("defnames.xlsx");
        ExcelWorksheet? ws1 = p.Workbook.Worksheets.Add("Sheet1");
        _ = p.Workbook.Worksheets.Add("Sheet2");

        ExcelNamedRange? name = ws1.Names.Add("Name2", ws1.Cells["B1:C5"]);
        Assert.AreEqual("Sheet1", name.Worksheet.Name);
        name.Address = "Sheet3!B2:C6";
        Assert.IsNull(name.Worksheet);
        Assert.AreEqual("Sheet3", name.WorkSheetName);
    }

    [TestMethod]
    public void Issue341()
    {
        using ExcelPackage? package = OpenTemplatePackage("Base_debug.xlsx");

        using (ExcelPackage? atomic_sheet_package = OpenTemplatePackage("Test_debug.xlsx"))
        {
            ExcelWorksheet? s = atomic_sheet_package.Workbook.Worksheets["Test3"];
            ExcelWorksheet? s_copy = package.Workbook.Worksheets.Add("Test3", s); // Exception on this line
            s_copy.Drawings[0].As.Chart.LineChart.Series[0].XSeries = "A1:a15";
            atomic_sheet_package.Save();
        }

        package.Save();
    }

    [TestMethod]
    public void Issue347_2()
    {
        using ExcelPackage? p = OpenTemplatePackage("i347.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue353()
    {
        using ExcelPackage? p = OpenTemplatePackage("HeaderFooterTest (1).xlsx");
        ExcelWorksheet worksheet = p.Workbook.Worksheets[0];
        Assert.IsFalse(worksheet.HeaderFooter.differentFirst);
        Assert.IsFalse(worksheet.HeaderFooter.differentOddEven);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue354()
    {
        using ExcelPackage? p = OpenTemplatePackage("i354.xlsx");
        ExcelWorksheet? ws1 = p.Workbook.Worksheets[0];
        ExcelWorksheet? ws2 = p.Workbook.Worksheets[2];
        ExcelPivotTable? pt = ws1.PivotTables.Add(ws1.Cells["A2"], ws2.Cells["A1:E3005"], "pt");
        ws2.Cells["B2"].Value = eDateGroupBy.Years;
        ws2.Cells["B3"].Value = eDateGroupBy.Months;
        _ = pt.ColumnFields.Add(pt.Fields[1]);
        _ = pt.RowFields.Add(pt.Fields[4]);
        pt.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void StyleIssueLibreOffice()
    {
        foreach (bool onColumns in new[] { true, false })
        {
            ExcelPackage? ep = new ExcelPackage();
            ExcelWorksheet? ws = ep.Workbook.Worksheets.Add("Test");

            // Header area (along with freezing the header in the view)
            ws.Cells[1, 1, 1, 8].Style.Font.Bold = true;

            for (int i = 1; i < 9; ++i)
            {
                ws.Cells[1, i].Value = $"Test {i}";
            }

            ws.View.FreezePanes(2, 1);

            if (onColumns)
            {
                // Set the horizontal alignment on the columns themselves
                ws.Column(3).Style.HorizontalAlignment = ws.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                ws.Column(5).Style.HorizontalAlignment = ws.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(7).Style.HorizontalAlignment = ws.Column(8).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }
            else
            {
                // Set the horizontal alignment on the cells of the header
                ws.Cells[1, 3, 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                ws.Cells[1, 5, 1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[1, 7, 1, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }

            for (int row = 2; row < 30; ++row)
            {
                for (int i = 1; i < 9; ++i)
                {
                    ws.Cells[row, i].Value = row % 2 == 0 ? ((8 * (row - 2)) + i).ToString() : $"Test {(8 * (row - 2)) + i}";
                }

                if (!onColumns)
                {
                    // Set the horizontal alignment on this row's cells
                    ws.Cells[row, 3, row, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 5, row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 7, row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }

            ws.Cells.AutoFitColumns(0);

            ep.SaveAs(new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                                $"AlignmentTest-On{(onColumns ? "Columns" : "Cells")}.xlsx")));
        }
    }

    [TestMethod]
    public void Issue382()
    {
        using ExcelPackage? p = OpenPackage("Issue382.xlsx", true);
        p.Workbook.Styles.NamedStyles[0].Style.Font.Size = 9;
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Cell Value";
        ws.Cells.AutoFitColumns();
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue381()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue381.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[1];
        Assert.AreEqual(2, ws.Drawings.Count);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void ReadIssue()
    {
        using ExcelPackage? p = OpenTemplatePackage("cf.xlsx");
    }

    [TestMethod]
    public void Issue407_1()
    {
        using ExcelPackage? p = OpenTemplatePackage("TestStyles_MoreCellStyleXfsThanCellXfs.xlsx");
        ExcelPackage? p2 = new ExcelPackage();
        _ = p2.Workbook.Worksheets.Add("Copied Style", p.Workbook.Worksheets["StylesTestSheet"]);

        SaveWorkbook("Issue407_1.xlsx", p2);
    }

    [TestMethod]
    public void Issue407_2()
    {
        using ExcelPackage? p = OpenTemplatePackage("TestStyles_MinimalWithNamedStyles.xlsx");
        ExcelPackage? p2 = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets["StylesTestSheet"];

        Assert.AreEqual("Normal", ws.Cells["A1"].StyleName);
        Assert.AreEqual("MyCustomCellStyle", ws.Cells["A2"].StyleName);
        Assert.AreEqual("Normal", ws.Cells["A3"].StyleName);
        Assert.AreEqual("MyCalculationStyle", ws.Cells["A4"].StyleName);

        Assert.AreEqual("Normal", ws.Cells["B1"].StyleName);
        Assert.AreEqual("MyBoldStyle1", ws.Cells["B2"].StyleName);
        Assert.AreEqual("MyBoldStyle2", ws.Cells["B3"].StyleName);
        ws = p2.Workbook.Worksheets.Add("Copied Style", p.Workbook.Worksheets["StylesTestSheet"]);
        Assert.AreEqual("Normal", ws.Cells["A1"].StyleName);
        Assert.AreEqual("MyCustomCellStyle", ws.Cells["A2"].StyleName);
        Assert.AreEqual("Normal", ws.Cells["A3"].StyleName);
        Assert.AreEqual("MyCalculationStyle", ws.Cells["A4"].StyleName);

        Assert.AreEqual("Normal", ws.Cells["B1"].StyleName);
        Assert.AreEqual("MyBoldStyle1", ws.Cells["B2"].StyleName);
        Assert.AreEqual("MyBoldStyle2", ws.Cells["B3"].StyleName);

        SaveWorkbook("Issue407_2.xlsx", p2);
    }

    [TestMethod]
    public void s185()
    {
        using ExcelPackage? p = OpenTemplatePackage("s185.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ExcelLineChart? chart = ws.Drawings[0] as ExcelLineChart;

        Assert.AreEqual(4887, chart.PlotArea.ChartTypes[1].Series[0].NumberOfItems);
    }

    [TestMethod]
    public void GetNamedRangeAddressAfterRowInsert()
    {
        using ExcelPackage? pck = OpenTemplatePackage("TestWbk_SingleNamedRange.xlsx");

        // Get the worksheet containing the named range
        ExcelWorksheet? ws = pck.Workbook.Worksheets["Sheet1"];

        // Get the named range
        ExcelNamedRange? namedRange = ws.Names["MyValues"];

        // Check that the named range exists with the expected address
        Assert.AreEqual("Sheet1!$A$1:$A$9", namedRange.FullAddress);
        Assert.AreEqual("Sheet1!$A$1:$A$9", namedRange.Address); // This line is currently failing

        // Insert a row in the middle of the range
        ws.InsertRow(5, 1);

        // Check that the named range's address has been correctly updated
        Assert.AreEqual("Sheet1!$A$1:$A$10", namedRange.FullAddress);
        Assert.AreEqual("Sheet1!$A$1:$A$10", namedRange.Address);
    }

    [TestMethod]
    public void Issue410()
    {
        using ExcelPackage? package = OpenTemplatePackage("test-in.xlsx");
        ExcelWorkbook? wb = package.Workbook;
        ExcelWorksheet? worksheet = wb.Worksheets.Add("Pivot Tables");
        ExcelTable? table = wb.Worksheets[0].Tables["Table1"];
        ExcelPivotTable pt = worksheet.PivotTables.Add(worksheet.Cells["A1"], table, "PT1");
        _ = pt.RowFields.Add(pt.Fields["ColC"]);
        _ = pt.DataFields.Add(pt.Fields["ColB"]);
        SaveWorkbook("test-out.xlsx", package);
    }

    [TestMethod]
    public void Issue415()
    {
        using ExcelPackage? package = OpenTemplatePackage("Issue415.xlsm");
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void Issue417()
    {
        using ExcelPackage? package = OpenTemplatePackage("Issue417.xlsx");
        ExcelWorksheet? ws = package.Workbook.Worksheets[0];
        Assert.AreEqual("0", ws.Cells["A1"].Text);
        Assert.AreEqual(null, ws.Cells["A2"].Text);
    }

    [TestMethod]
    public void Issue395()
    {
        using ExcelPackage? package = OpenTemplatePackage("Issue395.xlsx");
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void Issue418()
    {
        using ExcelPackage? p = OpenPackage("issue418.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Test");

        ExcelRange? mergetest = ws.Cells[2, 1, 2, 5];
        mergetest.IsRichText = true;
        mergetest.Merge = true;
        mergetest.Style.WrapText = true;

        ExcelRichText? t1 = mergetest.RichText.Add($"Text 1", true);
        t1.Size = 16;
        t1.Bold = true;

        ExcelRichText? t2 = mergetest.RichText.Add($"Text 2", true);
        t2.Size = 12;
        t2.Bold = false;

        ws.Row(2).Height = 50;

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void IssueHidden()
    {
        using ExcelPackage? p = OpenTemplatePackage("workbook.xlsx");
        p.Workbook.Worksheets[p.Workbook.Worksheets.Count - 2].Hidden = eWorkSheetHidden.Hidden;
        p.Workbook.Worksheets[p.Workbook.Worksheets.Count - 1].Hidden = eWorkSheetHidden.Hidden;
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue430()
    {
        using ExcelPackage? p = OpenTemplatePackage("issue430.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue435()
    {
        using ExcelPackage? p = OpenTemplatePackage("issue435.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void VbaIssueLoad()
    {
        using ExcelPackage? p = OpenTemplatePackage("PlantillaDefectivo-NotWorking.xlsm");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue440()
    {
        using ExcelPackage? p = OpenTemplatePackage("issue440.xlsx");
        ExcelWorkbook? wb = p.Workbook;
        ExcelWorksheet? worksheet = wb.Worksheets.Add("Pivot Tables");
        ExcelTable? table = wb.Worksheets[0].Tables["Table1"];
        ExcelPivotTable pt = worksheet.PivotTables.Add(worksheet.Cells["A1"], table, "PT1");
        _ = pt.RowFields.Add(pt.Fields["ColC"]);
        _ = pt.DataFields.Add(pt.Fields["ColB"]);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue441()
    {
        using ExcelPackage? pck = OpenPackage("issue441.xlsx", true);
        ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
        string? commentAddress = "B2";
        _ = wks.Comments.Add(wks.Cells[commentAddress], "This is a comment.", "author");
        wks.Cells[commentAddress].Value = "This cell contains a comment.";

        wks.Cells["B1:B3"].Insert(eShiftTypeInsert.Right);
        commentAddress = "C2";
        Assert.AreEqual(1, wks.Comments.Count);
        Assert.AreEqual("This is a comment.", wks.Comments[0].Text);
        Assert.AreEqual("This cell contains a comment.", wks.Cells[commentAddress].GetValue<string>());
        Assert.AreEqual(commentAddress, wks.Comments[0].Address);
        SaveAndCleanup(pck);
    }

    [TestMethod]
    public void Issue442()
    {
        using ExcelPackage? pck = OpenPackage("issue442.xlsx", true);

        // Add a sheet with data validation in cell B2
        ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
        IExcelDataValidationList? dataValidationList = wks.DataValidations.AddListValidation("B2");
        IList<string>? values = dataValidationList.Formula.Values;
        values.Add("Yes");
        values.Add("No");
        wks.Cells["B2"].Value = "Yes";

        // Confirm this was added in the right place
        Assert.AreEqual("Yes", wks.Cells["B2"].GetValue<string>());
        Assert.AreEqual("B2", wks.DataValidations[0].Address.Address);

        // Insert cells to shift this data validation to the right
        wks.Cells["B2:C5"].Insert(eShiftTypeInsert.Right);

        // Check the data validation has been moved to the right place
        Assert.AreEqual("Yes", wks.Cells["D2"].GetValue<string>());
        Assert.AreEqual("D2", wks.DataValidations[0].Address.Address);

        SaveAndCleanup(pck);
    }

    [TestMethod]
    public void s224()
    {
        using ExcelPackage? p = OpenTemplatePackage("s224.xltx");
        SaveWorkbook("s224.xlsx", p);
    }

    [TestMethod]
    public void InsertCellsNextToComment()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");

        // Create a comment in cell B2
        string? commentAddress = "B2";
        _ = wks.Comments.Add(wks.Cells[commentAddress], "This is a comment.", "author");
        wks.Cells[commentAddress].Value = "This cell contains a comment.";

        // Create another comment in cell B10
        string? commentAddress2 = "B10";
        _ = wks.Comments.Add(wks.Cells[commentAddress2], "This is another comment.", "author");
        wks.Cells[commentAddress2].Value = "This cell contains another comment.";

        // Insert cells so the first comment is now in C2
        wks.Cells["B1:B3"].Insert(eShiftTypeInsert.Right);
        commentAddress = "C2";

        // Check that both the cell value and the comment was correctly moved
        Assert.AreEqual(2, wks.Comments.Count);
        Assert.AreEqual("This is a comment.", wks.Comments[0].Text);
        Assert.AreEqual("This cell contains a comment.", wks.Cells[commentAddress].GetValue<string>());
        Assert.AreEqual(commentAddress, wks.Comments[0].Address);

        // Check that the second comment hasn't moved
        Assert.AreEqual("This is another comment.", wks.Comments[1].Text);
        Assert.AreEqual("This cell contains another comment.", wks.Cells[commentAddress2].GetValue<string>());
        Assert.AreEqual(commentAddress2, wks.Comments[1].Address);
    }

    [TestMethod]
    public void InsertRowsIntoTable_CheckFormulasWithColumnReferences()
    {
        using ExcelPackage? pck = new ExcelPackage();

        // Add a sheet, and a table with headers and a single row of data
        ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
        wks.Cells["B2:D3"].Value = new object[,] { { "Col1", "Col2", "Col3" }, { 1, 2, 3 } };
        _ = wks.Tables.Add(wks.Cells["B2:D3"], "Table1");

        // Add a SUM formula on the worksheet with a reference to a table column
        wks.Cells["C10"].Formula = "SUM(Table1[Col2])";
        wks.Cells["C10"].Calculate();

        Assert.AreEqual(2.0, wks.Cells["C3"].GetValue<double>());
        Assert.AreEqual("SUM(Table1[Col2])", wks.Cells["C10"].Formula);
        Assert.AreEqual(2.0, wks.Cells["C10"].GetValue<double>());

        // Insert 2 rows into the worksheet to extend the table
        wks.InsertRow(4, 2);

        // Check that the formula that was in C10 (now C12) still references the column
        Assert.AreEqual("SUM(Table1[Col2])", wks.Cells["C12"].Formula);
    }

    [TestMethod]
    public void CopyWorksheetWithDynamicArrayFormula()
    {
        using (ExcelPackage? p1 = OpenTemplatePackage("TestDynamicArrayFormula.xlsx"))
        {
            using ExcelPackage? p2 = new ExcelPackage();
            _ = p2.Workbook.Worksheets.Add("Sheet1", p1.Workbook.Worksheets["Sheet1"]);
            SaveWorkbook("DontCopyMetadataToNewWorkbook.xlsx", p2);
        }

        Assert.Inconclusive("Now try to open the file in Excel - does it tell you the file is corrupt?");
    }

    [TestMethod]
    public void VbaIssue()
    {
        using ExcelPackage? p = OpenTemplatePackage("Issue479.xlsm");
        _ = p.Workbook.Worksheets.Add("New Sheet");
        SaveAndCleanup(p);
    }

    public static readonly Dictionary<string, string> ASSET_FIELDS = new Dictionary<string, string>
    {
        { string.Empty, "Select one..." },
        { "APPRAISAL_DATE", "Appraisal Date" },
        { "APPRAISAL_AREA", "Appraisal Surface" },
        { "APPRAISAL_AREA_CCAA", "Appraisal Surface w/CCAA" },
        { "APPRAISAL_VALUE", "Appraisal Value" },
        { "AURA" + "." + "REFERENCE", "Aura ID" },
        { "BATHROOMS", "Bathroom" },
        { "BORROWER_ID", "Borrower ID" },
        { "AREA_CCAA", "Built surface w/CCAA" },
        { "CADASTRAL_REFERENCE", "Cadastral reference" },
        { "DEVELOPMENT" + "." + "CLIENT_GROUP_ID", "Client Development ID" },
        { "CLIENT_ID", "Client ID" },
        { "YEAR_OF_CONSTRUCTION", "Construction Year" },
        { "COUNTRY", "Country" },
        { "CROSSING_DOCKS", "Crossing Docks" },
        { "DATE_OF_DISQUALIFICATION", "Date of desaffection - Social Housing" },
        { "LEADER" + "." + "REFERENCE", "Dependency Reference" },
        { "DEVELOPMENT" + "." + "REFERENCE", "Development ID" },
        { "ADDRESS_DOOR", "Door" },
        { "DUPLEX", "Duplex" },
        { "ELEVATOR", "Elevator" },
        { "ADDRESS_FLOOR", "Floor" },
        { "FULL_ADDRESS", "Full Address" },
        { "IDUFIR", "IDUFIR" },
        { "ILLEGAL_SQUATTERS", "Illegal Squatters" },
        { "ORIENTATION", "Interior/Exterior" },
        { "LATITUDE", "Latitude" },
        { "LIEN", "Lien" },
        { "LOAN_ID", "Loan ID" },
        { "LONGITUDE", "Longitude" },
        { "MAINTENANCE_STATUS", "Maintenance Status" },
        { "MARKET_SHARE", "Market Share (%)" },
        { "LEGAL_MAXIMUM_VALUE", "Max. Value - Social Housing" },
        { "MAX_HEIGHT", "Maximum Height" },
        { "VPO_MODULE", "Module - Social Housing" },
        { "MUNICIPALITY", "Municipality" },
        { "NEGATIVE_COLD", "Negative Cold" },
        { "ADDRESS_NUMBER", "Number" },
        { "BORROWER", "Owner" },
        { "PARKINGS", "Parking" },
        { "PERIMETER", "Perimeter" },
        { "PLOT_AREA", "Plot Surface" },
        { "POSITIVE_COLD", "Positive Cold" },
        { "PROVINCE", "Province" },
        { "REFERENCE", "Reference ID" },
        { "REGISTRATION", "Registry" },
        { "REGISTRY_ID", "Registry ID" },
        { "REGISTRATION_NUMBER", "Registry Number" },
        { "AREA_REGISTRY", "Registry Surface" },
        { "AREA_CCAA_REGISTRY", "Registry Surface w/CCAA" },
        { "RENTED", "Rented" },
        { "REPEATED", "Repeated" },
        { "ROOMS", "Rooms" },
        { "SCOPE", "Scope" },
        { "SEA_VIEWS", "Sea Views" },
        { "MONTHLY_COMM_EXP_SQM", "Service Charges" },
        { "SMOKE_VENT", "Smoke Ventilation" },
        { "VPO", "Social Housing" },
        { "DEVELOPMENT" + "." + "PROPERTY_STATUS", "Status" },
        { "STOREROOMS", "Storage" },
        { "ADDRESS_NAME", "Street" },
        { "ASSET_SUBTYPE", "Sub-typology" },
        { "AREA", "Surface" },
        { "SWIMMING_POOL", "Swimming Pool" },
        { "TERRACE", "Terrace" },
        { "TERRACE_AREA", "Terrace Surface" },
        { "ACTIVITY", "Type of activity" },
        { "STATE", "Type of product" },
        { "ASSET_TYPE", "Typology" },
        { "USEFUL_AREA", "Useful Surface" },
        { "VALUATION_TYPE", "Valuation Type" },
        { "ZIP_CODE", "Zip Code" }
    };

    public class Error
    {
        public string TypeOfError { get; set; }

        public int Row { get; set; }

        public int Col { get; set; }

        public List<string> Messages { get; set; }
    }

    public class AssetField
    {
        public int Index { get; set; }

        public string Field { get; set; }
    }

    [TestMethod]
    public void Issue478()
    {
        int dataStartRow = 2;

        Error[]? errors =
            JsonConvert.DeserializeObject<Error[]>(
                                                   "[{\"typeOfError\":\"WARNING\",\"row\":4,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":20,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":35,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":47,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":57,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":60,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":90,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":131,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":136,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":138,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]},{\"typeOfError\":\"WARNING\",\"row\":139,\"col\":17,\"messages\":[\"The address is uncompleted. It can only get an approximate coordinates.\"]}]");

        AssetField[]? assetFields =
            JsonConvert
                .DeserializeObject<AssetField[]>("[{\"index\":1,\"field\":\"Reference\"},{\"index\":15,\"field\":\"ZipCode\"},{\"index\":16,\"field\":\"Municipality\"},{\"index\":17,\"field\":\"FullAddress\"}]");

        using (ExcelPackage? excelPackage = OpenTemplatePackage("issue478.xlsx"))
        {
            ExcelWorksheet? worksheet = excelPackage.Workbook.Worksheets["Avances"];
            ExcelCellAddress? end = worksheet.Dimension.End;

            // Add column of errors and warnings
            int startMessagesColumn = end.Column + 1;
            worksheet.InsertColumn(startMessagesColumn, 2);
            int warningColumn = startMessagesColumn + 1;
            worksheet.Cells[dataStartRow - 1, startMessagesColumn].Value = "Errors";
            worksheet.Cells[dataStartRow - 1, warningColumn].Value = "Warnings";

            foreach (Error? error in errors)
            {
                if (error.TypeOfError == "ERROR")
                {
                    //worksheet.Cells[error.Row - 1, errorColumn].Value += string.Join(" ", error.Messages.Select(w => string.Format("{0} {1}", ASSET_FIELDS.GetValueOrDefault(assetFields.Where(x => x.Index == error.Col).Select(x => x.Field).FirstOrDefault()), w)));
                }
                else
                {
                    //worksheet.Cells[error.Row - 1, warningColumn].Value += string.Join(" ", error.Messages.Select(w => string.Format("{0} {1}", ASSET_FIELDS.GetValueOrDefault(assetFields.Where(x => x.Index == error.Col).Select(x => x.Field).FirstOrDefault()), w)));
                }
            }

            // Remove distinct columns from "Reference"
            int colFieldReference = assetFields.Where(x => x.Field == "REFERENCE").Select(x => x.Index).FirstOrDefault();
            worksheet.Cells[1, colFieldReference + 1].Value = "Reference";

            int deletedColumns = 0;

            for (int i = 1; i <= end.Column; i++)
            {
                if (colFieldReference + 1 != i && startMessagesColumn != i && warningColumn != i)
                {
                    worksheet.DeleteColumn(i - deletedColumns);
                    deletedColumns++;
                }
            }

            // Remove rows that do not contain errors
            int deletedRows = 0;

            for (int i = 1; i <= end.Row; i++)
            {
                if (i < dataStartRow - 1 || (i >= dataStartRow && errors.All(w => w.Row - 1 != i)))
                {
                    worksheet.DeleteRow(i - deletedRows);
                    deletedRows++;
                }
            }

            SaveAndCleanup(excelPackage);
        }

        ;
    }

    [TestMethod]
    public void TestColumnWidthsAfterDeletingColumn()
    {
        using ExcelPackage? pck = OpenTemplatePackage("Issue480.xlsx");

        // Get the worksheet where columns 3-5 have a width of around 18
        ExcelWorksheet? wks = pck.Workbook.Worksheets["Sheet1"];

        // Check the width of column 5
        Assert.AreEqual(18.77734375, wks.Column(5).Width, 1E-5);

        // Delete column 4
        wks.DeleteColumn(4, 3);

        // Check width of column 5 (now 4) hasn't changed
        Assert.AreEqual(18.77734375, wks.Column(3).Width, 1E-5);

        //Assert.AreEqual(18.77734375, wks.Column(4).Width, 1E-5);
    }

    [TestMethod]
    public void Issue484_InsertRowCalculatedColumnFormula()
    {
        using ExcelPackage? p = new ExcelPackage();

        // Create some worksheets
        ExcelWorksheet? ws1 = p.Workbook.Worksheets.Add("Sheet1");

        // Create some tables with calculated column formulas
        ExcelTable? tbl1 = ws1.Tables.Add(ws1.Cells["A11:C12"], "Table1");
        tbl1.Columns[2].CalculatedColumnFormula = "A12+B12";

        ExcelTable? tbl2 = ws1.Tables.Add(ws1.Cells["E11:G12"], "Table2");
        tbl2.Columns[2].CalculatedColumnFormula = "A12+F12";

        // Check the formulas have been set correctly
        Assert.AreEqual("A12+B12", ws1.Cells["C12"].Formula);
        Assert.AreEqual("A12+F12", ws1.Cells["G12"].Formula);
        Assert.AreEqual("A12+B12", tbl1.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("A12+F12", tbl2.Columns["Column3"].CalculatedColumnFormula);

        // Delete two rows above the tables
        ws1.DeleteRow(5, 2);

        // Check the formulas were updated
        Assert.AreEqual("A10+B10", ws1.Cells["C10"].Formula);
        Assert.AreEqual("A10+F10", ws1.Cells["G10"].Formula);
        Assert.AreEqual("A10+B10", tbl1.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("A10+F10", tbl2.Columns[2].CalculatedColumnFormula);
    }

    [TestMethod]
    public void Issue484_DeleteRowCalculatedColumnFormula()
    {
        using ExcelPackage? p = new ExcelPackage();

        // Create some worksheets
        ExcelWorksheet? ws1 = p.Workbook.Worksheets.Add("Sheet1");

        // Create some tables with calculated column formulas
        ExcelTable? tbl1 = ws1.Tables.Add(ws1.Cells["A11:C12"], "Table1");
        tbl1.Columns[2].CalculatedColumnFormula = "A12+B12";

        ExcelTable? tbl2 = ws1.Tables.Add(ws1.Cells["E11:G12"], "Table2");
        tbl2.Columns[2].CalculatedColumnFormula = "A12+F12";

        // Check the formulas have been set correctly
        Assert.AreEqual("A12+B12", ws1.Cells["C12"].Formula);
        Assert.AreEqual("A12+F12", ws1.Cells["G12"].Formula);
        Assert.AreEqual("A12+B12", tbl1.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("A12+F12", tbl2.Columns["Column3"].CalculatedColumnFormula);

        // Delete two rows above the tables
        ws1.DeleteRow(5, 2);

        // Check the formulas were updated
        Assert.AreEqual("A10+B10", ws1.Cells["C10"].Formula);
        Assert.AreEqual("A10+F10", ws1.Cells["G10"].Formula);
        Assert.AreEqual("A10+B10", tbl1.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("A10+F10", tbl2.Columns[2].CalculatedColumnFormula);
    }

    [TestMethod]
    public void FreezeTemplate()
    {
        using ExcelPackage? p = OpenTemplatePackage("freeze.xlsx");

        // Get the worksheet where columns 3-5 have a width of around 18
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ws.View.FreezePanes(40, 5);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void CopyWorksheetWithBlipFillObjects()
    {
        using ExcelPackage? p1 = OpenTemplatePackage("BlipFills.xlsx");
        ExcelWorksheet? ws = p1.Workbook.Worksheets[0];
        ExcelWorksheet? wsCopy = p1.Workbook.Worksheets.Add("Copy", p1.Workbook.Worksheets[0]);

        ws.Cells["G4"].Copy(wsCopy.Cells["F20"]);
        SaveAndCleanup(p1);
    }

    [TestMethod]
    public void Issue519()
    {
        using ExcelPackage? package = OpenPackage("I519.xlsx", true);

        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

        worksheet.Cells[3, 1].Value = "test";

        ExcelControlCheckBox? ctrl = worksheet.Drawings.AddCheckBoxControl("test");
        ctrl.SetPosition(10, 10);
        ctrl.Checked = eCheckState.Checked; // creates valid XLSX file

        //ctrl.Checked = OfficeOpenXml.Drawing.Controls.eCheckState.Mixed; // creates valid XLSX file
        //ctrl.Checked = OfficeOpenXml.Drawing.Controls.eCheckState.Unchecked; // creates invalid XLSX file

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void Issue520_2()
    {
        using ExcelPackage? p = OpenPackage("i520.xlsx", true);
        ExcelWorksheet? sheet = p.Workbook.Worksheets.Add("17Columns");

        var tableData = Enumerable.Range(1, 10)
                                  .Select(_ => new
                                  {
                                      C01 = 1,
                                      C02 = 2,
                                      C03 = 3,
                                      C04 = 4,
                                      C05 = 5,
                                      C06 = 6,
                                      C07 = 7,
                                      C08 = 8,
                                      C09 = 9,
                                      C10 = 10,
                                      C11 = 11,
                                      C12 = 12,
                                      C13 = 13,
                                      C14 = 14,
                                      C15 = 15,
                                      C16 = 16,
                                      C17 = 17
                                  })
                                  .ToArray();

        ExcelRangeBase? table = sheet.Cells[1, 1].LoadFromCollection(tableData, true, TableStyles.Light1);
        table.AutoFitColumns();

        sheet = p.Workbook.Worksheets.Add("16Columns");

        var tableData2 = Enumerable.Range(1, 10)
                                   .Select(_ => new
                                   {
                                       C01 = 1,
                                       C02 = 2,
                                       C03 = 3,
                                       C04 = 4,
                                       C05 = 5,
                                       C06 = 6,
                                       C07 = 7,
                                       C08 = 8,
                                       C09 = 9,
                                       C10 = 10,
                                       C11 = 11,
                                       C12 = 12,
                                       C13 = 13,
                                       C14 = 14,
                                       C15 = 15,
                                       C16 = 16
                                   })
                                   .ToArray();

        table = sheet.Cells[1, 1].LoadFromCollection(tableData2, true, TableStyles.Light1);
        table.AutoFitColumns();

        sheet = p.Workbook.Worksheets.Add("18Columns");

        var tableData3 = Enumerable.Range(1, 10)
                                   .Select(_ => new
                                   {
                                       C01 = 1,
                                       C02 = 2,
                                       C03 = 3,
                                       C04 = 4,
                                       C05 = 5,
                                       C06 = 6,
                                       C07 = 7,
                                       C08 = 8,
                                       C09 = 9,
                                       C10 = 10,
                                       C11 = 11,
                                       C12 = 12,
                                       C13 = 13,
                                       C14 = 14,
                                       C15 = 15,
                                       C16 = 16,
                                       C17 = 17,
                                       C18 = 18
                                   })
                                   .ToArray();

        table = sheet.Cells[1, 1].LoadFromCollection(tableData3, true, TableStyles.Light1);
        table.AutoFitColumns();
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue522()
    {
        using ExcelPackage? package = OpenPackage("I22.xlsx", true);

        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

        worksheet.Cells[1, 1].Value = -1234;
        worksheet.Cells[1, 1].Style.Numberformat.Format = "#.##0\"*\";(#.##0)\"*\"";

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void IssueNamedRanges()
    {
        using ExcelPackage? package = OpenTemplatePackage("ORRange23 Problem.xlsx");

        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

        worksheet.Cells[1, 1].Value = -1234;
        worksheet.Cells[1, 1].Style.Numberformat.Format = "#.##0\"*\";(#.##0)\"*\"";

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void DvcfCopy()
    {
        using ExcelPackage? p = OpenTemplatePackage("i527.xlsm");

        // Fails when data validation is set
        // Fails when conditional formatting is set.
        ExcelRange? copyFrom1 = p.Workbook.Worksheets["CopyFrom"].Cells["A1:BR23"];
        ExcelRange? copyTo1 = p.Workbook.Worksheets["CopyTo"].Cells["A:XFD"];
        copyFrom1.Copy(copyTo1);

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s268()
    {
        using ExcelPackage? p = OpenTemplatePackage("s268.xlsx");
        ExcelWorksheet? s3 = p.Workbook.Worksheets["s3"];

        s3.InsertRow(1, 1);
        s3.InsertRow(1, 1);
        s3.InsertRow(1, 1);
        s3.InsertRow(1, 1);
        s3.InsertRow(1, 1);
        s3.InsertRow(1, 1);
        s3.InsertRow(1, 1);

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue538()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Sheet1");
        _ = package.Workbook.Worksheets.Add("Sheet2");
        IExcelDataValidationList? validation = sheet.DataValidations.AddListValidation("A1 B1");
        validation.Formula.ExcelFormula = "Sheet2!$A$7:$A$12"; // throws exception "Multiple addresses may not be commaseparated, use space instead"
    }

    [TestMethod]
    public void s272()
    {
        using ExcelPackage? p = OpenTemplatePackage("RadioButton.xlsm");

        if (p.Workbook.VbaProject == null)
        {
            p.Workbook.CreateVBAProject();
        }

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s277()
    {
        using ExcelPackage? p = OpenTemplatePackage("s277.xlsx");

        foreach (ExcelWorksheet? ws in p.Workbook.Worksheets)
        {
            ws.Drawings.Clear();
        }
    }

    [TestMethod]
    public void s279()
    {
        using ExcelPackage? p = OpenTemplatePackage("s279.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ws.Cells["C3"].Value = "Test";
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I546()
    {
        using ExcelPackage? excelPackage = OpenTemplatePackage("b.xlsx");
        ExcelWorksheet? ws = excelPackage.Workbook.Worksheets[0];
        ExcelRange? cell = ws.Cells["A2"];
        object? value1 = cell.Value;
        Console.WriteLine($"value1: {value1}");

        ExcelExternalLinksCollection? externalLinks = excelPackage.Workbook.ExternalLinks;
        ExcelExternalWorkbook? externalWorkbook = externalLinks[0].As.ExternalWorkbook;
        _ = externalWorkbook.Load();

        ws.ClearFormulaValues();
        ws.Calculate(); // "Circular reference occurred at A2" exception is thrown here

        object? value2 = cell.Value;
        Console.WriteLine($"value2: {value2}");
    }

    [TestMethod]
    public void I548()
    {
        using ExcelPackage? p = OpenTemplatePackage("09-145.xlsx");
        ExcelWorksheet? wsCopy = p.Workbook.Worksheets["Sheet3"];
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("tmpCopy");

        //copy in the same o in another workbook, same issue
        wsCopy.Cells["C1:AB55"].Copy(ws.Cells["C1"], ExcelRangeCopyOptionFlags.ExcludeFormulas);

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I552()
    {
        using (ExcelPackage? package = OpenTemplatePackage("I552-2.xlsx"))
        {
            ExcelWorksheet? worksheet = package.Workbook.Worksheets[0];
            worksheet.InsertRow(2, 1);
            worksheet.Cells[1, 1, 1, 10].Copy(worksheet.Cells[2, 1, 2, 10]);

            SaveAndCleanup(package);
        }

        using (ExcelPackage? package = OpenPackage("I552-2.xlsx"))
        {
            ExcelWorksheet? worksheet = package.Workbook.Worksheets[0];
            worksheet.InsertRow(2, 1);
            worksheet.Cells[1, 1, 1, 10].Copy(worksheet.Cells[2, 1, 2, 10]);

            SaveAndCleanup(package);
        }
    }

    [TestMethod]
    public void s285()
    {
        using ExcelPackage? package = OpenTemplatePackage("s285.xlsx");
        ExcelWorksheet? worksheet = package.Workbook.Worksheets[0];
        worksheet.SetValue(3, 3, "Test");
        ExcelNamedStyleXml? ns = package.Workbook.Styles.CreateNamedStyle("Normal");
        ns.BuildInId = 0;
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void i566()
    {
        using ExcelPackage? package = OpenPackage("i566.xlsx", true);
        _ = package.Workbook.Worksheets.Add("Sheet 1");
        ExcelWorksheet? ws = package.Workbook.Worksheets["Sheet 1"];
        ws.SetValue(3, 3, "Test");
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void i583()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.SetValue(1048576, 1, 1);
        Assert.AreEqual("A1048576", ws.Dimension.Address);
    }

    [TestMethod]
    public void i567()
    {
        using ExcelPackage? package = OpenTemplatePackage("i567.xlsx");
        ExcelWorksheet? wsSource = package.Workbook.Worksheets["Detail"];

        List<object[]>? dataCollection = new List<object[]>()
        {
            new object[] { "Driver 1", 1, 2, "Fleet 1", "Manager 1", 3, true, 0, 5, 0 },
            new object[] { "Driver 2", 3, 4, "Fleet 2", "Manager 2", 5, true, 0, 8, 0 }
        };

        wsSource.Cells["A1"].Value = null;

        //code to load a collection to the spreadsheet. very nice
        _ = wsSource.Cells["A2"].LoadFromArrays(dataCollection);

        foreach (ExcelWorksheet? ws in package.Workbook.Worksheets)
        {
            foreach (ExcelPivotTable? pt in ws.PivotTables)
            {
                pt.CacheDefinition.SourceRange = wsSource.Cells["A1:J3"];
            }
        }

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void i574()
    {
        using ExcelPackage? package = OpenTemplatePackage("i574.xlsx");

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void LoadFontSize() => FontSize.LoadAllFontsFromResource();

    [TestMethod]
    public void PiechartWithHorizontalSource()
    {
        using ExcelPackage? p = OpenPackage("piechartHorizontal.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("PieVertical");
        ws.SetValue("A1", "C1");
        ws.SetValue("A2", "C2");
        ws.SetValue("A3", "C3");
        ws.SetValue("B1", 15);
        ws.SetValue("B2", 45);
        ws.SetValue("B3", 40);

        ExcelPieChart? chart = ws.Drawings.AddPieChart("Pie1", ePieChartType.Pie);
        chart.VaryColors = true;
        _ = chart.Series.Add("B1:B3", "A1:A3");
        chart.StyleManager.SetChartStyle(ePresetChartStyle.PieChartStyle1);

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void PiechartWithVerticalSource()
    {
        using ExcelPackage? p = OpenPackage("piechartvertical.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("PieVertical");
        ws.SetValue("A1", "C1");
        ws.SetValue("B1", "C2");
        ws.SetValue("C1", "C3");
        ws.SetValue("A2", 15);
        ws.SetValue("B2", 45);
        ws.SetValue("C2", 40);

        ExcelPieChart? chart = ws.Drawings.AddPieChart("Pie1", ePieChartType.Pie);
        chart.VaryColors = true;
        _ = chart.Series.Add("A2:C2", "A1:C1");
        chart.StyleManager.SetChartStyle(ePresetChartStyle.PieChartStyle1);

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void CheckEnvironment() => _ = Graphics.FromHwnd(IntPtr.Zero);

    [TestMethod]
    public void Issue592()
    {
        using ExcelPackage? p = OpenTemplatePackage("I592.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I594()
    {
        using ExcelPackage? p = OpenTemplatePackage("i594.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[1];
        ExcelTable? tbl = ws.Tables[0];
        _ = tbl.AddRow(2);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s302()
    {
        using ExcelPackage? p = OpenTemplatePackage("SampleData.xlsx");
        ExcelWorksheet? worksheet = p.Workbook.Worksheets[2];
        ExcelBarChart chart = worksheet.Drawings.AddBarChart("NewBarChart", eBarChartType.BarClustered);

        chart.SetPosition(32, 0, 1, 0);

        chart.SetSize(785, 320);

        chart.RoundedCorners = false;

        chart.Border.Fill.Color = Color.Gray;

        chart.Legend.Position = eLegendPosition.Bottom;

        ExcelBarChartSerie eventS1Serie = chart.Series.Add("D9:D12", "B9:B12");

        eventS1Serie.Header = "STATISTIQUES COMPARATIVES";

        _ = chart.Series.Add("H9:H12", "B9:B12");

        chart.StyleManager.SetChartStyle(ePresetChartStyle.BarChartStyle5, ePresetChartColors.MonochromaticPalette5);

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I596()
    {
        using ExcelPackage? p = OpenTemplatePackage("I596.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I606()
    {
        using ExcelPackage? p = OpenTemplatePackage("i606.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s312()
    {
        using ExcelPackage? p = OpenTemplatePackage("richtext.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        Type? t = ws.Cells["C2"].RichText.GetType();
        ;
        PropertyInfo? prop = t.GetProperty("TopNode", BindingFlags.GetProperty | BindingFlags.NonPublic | BindingFlags.Instance);
        _ = prop.GetValue(ws.Cells["C2"].RichText);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue308()
    {
        using ExcelPackage package = new ExcelPackage();
        package.Compatibility.IsWorksheets1Based = false;

        ExcelWorkbook? wb = package.Workbook;
        _ = wb.Worksheets.Add("A");
        _ = wb.Worksheets.Add("B");

        ExcelWorksheet sheetA = wb.Worksheets["A"];
        ExcelWorksheet sheetB = wb.Worksheets["B"];

        string? qsUndefined = QStr("Undefined");
        string? qsEmpty = QStr("");
        string? qsOk = QStr("OK");
        string? qsError = QStr("ERROR");
        string? qsES4 = QStr("ES4");
        string? qsES5 = QStr("ES5");

        string? errorFormula = $"IF(OR(LEFT(RC4,9)={qsUndefined},AND(OR(RC5={qsES4},RC5={qsES5}),RC8={qsEmpty}),RC5={qsEmpty}),{qsError},{qsOk})";

        sheetB.Cells[1, 4].Value = "ABC";
        sheetB.Cells[1, 5].Value = "ES1";
        sheetB.Cells[1, 8].Value = "ES1";
        sheetB.Cells[1, 10].FormulaR1C1 = errorFormula;

        wb.Calculate();

        Console.WriteLine($"Sht B Value = {sheetB.Cells[1, 10].Value} FormulaR1c1={sheetB.Cells[1, 10].FormulaR1C1}");
        Console.WriteLine($"Sht A Value = {sheetA.Cells[1, 10].Value} FormulaR1c1={sheetA.Cells[1, 10].FormulaR1C1}");

        Assert.AreEqual("OK", sheetB.Cells["J1"].Value);

        // package.SaveAs(@"c:\temp\eppTest306.xlsx");
    }

    static string QStr(string s)
    {
        char quotechar = '\"';

        return $"{quotechar}{s}{quotechar}";
    }

    [TestMethod]
    public void s314()
    {
        using ExcelPackage? p = OpenTemplatePackage("SlicerIssue.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ExcelPivotTable? pt = ws.PivotTables[0];
        ExcelWorksheet? wsTable = p.Workbook.Worksheets[1];
        ExcelTable? tbl = wsTable.Tables[0];
        wsTable.Cells["E9"].Value = "New Value";
        wsTable.Cells["E9"].Value = "Stockholm";
        wsTable.Cells["F11"].Value = "Test";
        _ = tbl.AddRow(11);
        tbl.Range.Offset(1, 0, tbl.Range.Rows - 1, tbl.Range.Columns).Copy(tbl.Range.Offset(11, 0));
        Assert.IsNotNull(pt.Fields[2].Items.Count);
        Assert.IsNotNull(pt.Fields[10].Items.Count);
        Assert.IsNotNull(pt.Fields[1].Items.Count);

        for (int r = 12; r < 23; r++)
        {
            wsTable.Cells[r, 1].Value += "-" + r;
            wsTable.Cells[r, 2].Value = r;
            wsTable.Cells[r, 3].Value += "-" + r;
        }

        wsTable.Cells[15, 1].Value = null;
        wsTable.Cells[14, 2].Value = null;
        wsTable.Cells[13, 3].Value = null;

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i620()
    {
        using ExcelPackage? p = OpenTemplatePackage("i621.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ws.DeleteColumn(1, 3);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i609()
    {
        using ExcelPackage? p = OpenTemplatePackage("i609.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void HeaderFooterWithInitWhiteSpace()
    {
        using ExcelPackage? p = OpenPackage("i631.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.HeaderFooter.EvenFooter.RightAlignedText = "  Row1\r\nRow 2 ";
        ws.HeaderFooter.OddFooter.RightAlignedText = "\r\nRow1\r\nRow 2\r\n";
        ws.HeaderFooter.EvenHeader.LeftAlignedText = "\tRow1\r\nRow 2";
        ws.HeaderFooter.OddHeader.LeftAlignedText = " Row1\r\nRow 2";
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void CopyDxfs()
    {
        using ExcelPackage? p = OpenTemplatePackage("Input Sheet.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        using ExcelPackage? p2 = OpenPackage("CopyDxfs.xlsx", true);
        _ = p2.Workbook.Worksheets.Add("Sheet1", ws);
        SaveAndCleanup(p2);
    }

    [TestMethod]
    public void i654()
    {
        using ExcelPackage? p = OpenTemplatePackage("i654.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];

        for (int i = 0; i < 10; i++)
        {
            ws.InsertRow(5, 1);
        }

        SaveWorkbook("i654-saved.xlsx", p);
    }

    [TestMethod]
    public void I653()
    {
        using ExcelPackage? p = OpenTemplatePackage("i653.xlsx");
        ExcelWorksheet sheet = p.Workbook.Worksheets[0];

        for (int i = 3; i < 1003; i++)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            sheet.InsertRow(i, 1);
            sw.Stop();
            sheet.Cells[i, 2].Value = sw.ElapsedTicks;
        }

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I665()
    {
        using ExcelPackage? p1 = OpenTemplatePackage("Source.xlsx");
        ExcelWorksheet sheet = p1.Workbook.Worksheets[0];
        using ExcelPackage? p2 = OpenTemplatePackage("VbaCopy.xlsm");
        _ = p2.Workbook.Worksheets.Add("sheet2", sheet);
        SaveWorkbook("i665.xlsm", p2);
    }

    [TestMethod]
    public void I667()
    {
        using ExcelPackage? p = OpenTemplatePackage("I667.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];

        int lastDataRowIdx = 10;

        IEnumerable<ExcelCellAddress>? notCopyCols = ws.Cells[1, 1, 1, 100]
                                                       .Where(x => x.Value != null
                                                                   && x.Value.ToString()
                                                                       .Trim()
                                                                       .Replace(" ", string.Empty)
                                                                       .IndexOf("#NotCopy", StringComparison.InvariantCultureIgnoreCase)
                                                                   >= 0)
                                                       .Select(x => x.End);

        foreach (ExcelCellAddress? col in notCopyCols)
        {
            ws.Cells[lastDataRowIdx + 1, col.Column].Clear();
        }

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I673()
    {
        using ExcelPackage? p = OpenTemplatePackage("I673.xlsm");
        _ = p.Workbook.Worksheets.Copy(p.Workbook.Worksheets.First().Name, "copied sheet");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void SaveDefinedName()
    {
        using ExcelPackage? p = OpenTemplatePackage("SaveIssueName.xlsm");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I676()
    {
        using ExcelPackage? p = OpenTemplatePackage("i676.xlsm");
        ExcelWorksheet? ws = p.Workbook.Worksheets["4"];
        ws.Cells["P10"].Value = 10;
        p.Workbook.VbaProject.Remove();
        SaveWorkbook("i676.xlsx", p);
    }

    [TestMethod]
    public void s350()
    {
        using ExcelPackage? p = OpenTemplatePackage("s350.xlsm");
        SaveWorkbook("s350.xlsm", p);
    }

    [TestMethod]
    public void s351()
    {
        using ExcelPackage? p = OpenPackage("s351.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("sheet1");
        ExcelControlGroupBox? grpBox = ws.Drawings.AddGroupBoxControl("GB_Deliverables");

        int optionBoxWidth = 300;
        int optionBoxHeight = 27;

        int OptionBoxRowOffsetPixels = 10;

        grpBox.SetPosition(Row: 25, RowOffsetPixels: 4, Column: 2, ColumnOffsetPixels: 1);
        grpBox.SetSize(optionBoxWidth * 3, optionBoxHeight * 5);
        grpBox.Text = "";

        ExcelControlCheckBox? r1c1 = ws.Drawings.AddCheckBoxControl("General_Technimark_Submission");
        r1c1.Text = "General Technimark Submission";
        r1c1.SetPosition(25, OptionBoxRowOffsetPixels, 2, 5);
        r1c1.SetSize(optionBoxWidth, optionBoxHeight);

        //Still need to Create function to set cell to true/false based on IDArray[] values
        ws.Cells["J26"].Value = true;
        r1c1.LinkedCell = ws.Cells["J26"];

        ExcelControlCheckBox? r1c2 = ws.Drawings.AddCheckBoxControl("Project_Schedule");
        r1c2.Text = "Project Schedule";
        r1c2.SetPosition(25, OptionBoxRowOffsetPixels, 3, 5);
        r1c2.SetSize(optionBoxWidth, optionBoxHeight);
        r1c2.LinkedCell = ws.Cells["K26"];

        ExcelControlCheckBox? r1c3 = ws.Drawings.AddCheckBoxControl("Customer_Company_Questionnaire");
        r1c3.Text = "Customer Company Questionnaire";
        r1c3.SetPosition(25, OptionBoxRowOffsetPixels, 4, 5);
        r1c3.SetSize(optionBoxWidth, optionBoxHeight);
        r1c3.LinkedCell = ws.Cells["L26"];

        //*********************************************************
        ExcelControlCheckBox? r2c1 = ws.Drawings.AddCheckBoxControl("Customer_Provided_Bid_Sheets");
        r2c1.Text = "Customer Provided Bid Sheets";
        r2c1.SetPosition(26, OptionBoxRowOffsetPixels, 2, 5);
        r2c1.SetSize(optionBoxWidth, optionBoxHeight);
        r2c1.LinkedCell = ws.Cells["J27"];

        ExcelControlCheckBox? r2c2 = ws.Drawings.AddCheckBoxControl("PowerPoint_Presentation");
        r2c2.Text = "PowerPoint Presentation";
        r2c2.SetPosition(26, OptionBoxRowOffsetPixels, 3, 5);
        r2c2.SetSize(optionBoxWidth, optionBoxHeight);
        r2c2.LinkedCell = ws.Cells["K27"];

        ExcelControlCheckBox? r2c3 = ws.Drawings.AddCheckBoxControl("Comprehensive_Tooling_Automation_Quotes");
        r2c3.Text = "Comprehensive Tooling/Automation Quotes";
        r2c3.SetPosition(26, OptionBoxRowOffsetPixels, 4, 5);
        r2c3.SetSize(optionBoxWidth, optionBoxHeight);
        r2c3.LinkedCell = ws.Cells["L27"];

        //*********************************************************
        ExcelControlCheckBox? r3c1 = ws.Drawings.AddCheckBoxControl("Validation_Quote");
        r3c1.Text = "Validation Quote";
        r3c1.SetPosition(27, OptionBoxRowOffsetPixels, 2, 5);
        r3c1.SetSize(optionBoxWidth, optionBoxHeight);
        r3c1.LinkedCell = ws.Cells["J28"];

        ExcelControlCheckBox? r3c2 = ws.Drawings.AddCheckBoxControl("Process_Flow_Diagrams");
        r3c2.Text = "Process Flow Diagrams";
        r3c2.SetPosition(27, OptionBoxRowOffsetPixels, 3, 5);
        r3c2.SetSize(optionBoxWidth, optionBoxHeight);
        r3c2.LinkedCell = ws.Cells["K28"];

        ExcelControlCheckBox? r3c3 = ws.Drawings.AddCheckBoxControl("DFM_DesignForManufacturability");
        r3c3.Text = "DFM - Design For Manufacturability";
        r3c3.SetPosition(27, OptionBoxRowOffsetPixels, 4, 5);
        r3c3.SetSize(optionBoxWidth, optionBoxHeight);
        r3c3.LinkedCell = ws.Cells["L28"];

        //*********************************************************
        ExcelControlCheckBox? r4c1 = ws.Drawings.AddCheckBoxControl("Plant_Layout");
        r4c1.Text = "Plant Layout";
        r4c1.SetPosition(28, OptionBoxRowOffsetPixels, 2, 5);
        r4c1.SetSize(optionBoxWidth, optionBoxHeight);
        r4c1.LinkedCell = ws.Cells["J29"];

        ExcelControlCheckBox? r4c2 = ws.Drawings.AddCheckBoxControl("Moldflow");
        r4c2.Text = "Moldflow";
        r4c2.SetPosition(28, OptionBoxRowOffsetPixels, 3, 5);
        r4c2.SetSize(optionBoxWidth, optionBoxHeight);
        r4c2.LinkedCell = ws.Cells["K29"];

        ExcelControlCheckBox? r4c3 = ws.Drawings.AddCheckBoxControl("Development_Prototype_Proposal");
        r4c3.Text = "Development/Prototype Proposal";
        r4c3.SetPosition(28, OptionBoxRowOffsetPixels, 4, 5);
        r4c3.SetSize(optionBoxWidth, optionBoxHeight);
        r4c3.LinkedCell = ws.Cells["L29"];

        //*********************************************************
        ExcelControlCheckBox? r5c1 = ws.Drawings.AddCheckBoxControl("Amortization_Schedule");
        r5c1.Text = "Amortization Schedule";
        r5c1.SetPosition(29, OptionBoxRowOffsetPixels, 2, 5);
        r5c1.SetSize(optionBoxWidth, optionBoxHeight);
        r5c1.LinkedCell = ws.Cells["J30"];

        ExcelControlCheckBox? r5c2 = ws.Drawings.AddCheckBoxControl("Risk_Assessment");
        r5c2.Text = "Risk Assessment";
        r5c2.SetPosition(29, OptionBoxRowOffsetPixels, 3, 5);
        r5c2.SetSize(optionBoxWidth, optionBoxHeight);
        r5c2.LinkedCell = ws.Cells["K30"];

        ExcelControlCheckBox? r5c3 = ws.Drawings.AddCheckBoxControl("Total_CostOfOwnershipAnalysis");
        r5c3.Text = "Total Cost of Ownership Analysis";
        r5c3.SetPosition(29, OptionBoxRowOffsetPixels, 4, 5);
        r5c3.SetSize(optionBoxWidth, optionBoxHeight);
        r5c3.LinkedCell = ws.Cells["L30"];

        //*********************************************************
        ExcelControlCheckBox? r6c1 = ws.Drawings.AddCheckBoxControl("Organization Chart");
        r6c1.Text = "Organization Chart";
        r6c1.SetPosition(30, OptionBoxRowOffsetPixels, 2, 5);
        r6c1.SetSize(optionBoxWidth, optionBoxHeight);
        r6c1.LinkedCell = ws.Cells["J31"];

        //*********************************************************
        _ = grpBox.Group(r1c1, r1c2, r1c3, r2c1, r2c2, r2c3, r3c1, r3c2, r3c3, r4c1, r4c2, r4c3, r5c1, r5c2, r5c3, r6c1);

        ws.Row(32).Height = 30; //Add extra space between Strategy section
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i681()
    {
        using ExcelPackage? p = OpenTemplatePackage("i681.xlsx");
        p.Workbook.Calculate();

        ExcelWorksheet? ws = p.Workbook.Worksheets[1];
        Assert.AreEqual(400D, ws.Cells["B118"].Value);
    }

    [TestMethod]
    public void SumWithDoubleWorksheetRefs()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? wsA = p.Workbook.Worksheets.Add("a");
        ExcelWorksheet? wsB = p.Workbook.Worksheets.Add("b");
        wsA.Cells["A1"].Value = 1;
        wsA.Cells["A2"].Value = 2;
        wsA.Cells["A3"].Value = 3;
        wsB.Cells["A4"].Formula = "sum(a!a1:'a'!A3)";

        wsB.Calculate();
        Assert.AreEqual(6D, wsB.GetValue(4, 1));
    }

    [TestMethod]
    public void AddressWithDoubleWorksheetRefs()
    {
        ExcelAddressBase? a = new ExcelAddressBase("a!a1:'a'!A3");

        Assert.AreEqual(1, a._fromRow);
        Assert.AreEqual(3, a._toRow);
        Assert.AreEqual(1, a._fromCol);
        Assert.AreEqual(1, a._toCol);
    }

    [TestMethod]
    public void IsError_CellReference_StringLiteral()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
        sheet1.Cells["B2"].Value = ExcelErrorValue.Create(eErrorType.Value);
        sheet1.Cells["C3"].Formula = "ISERROR(B2)";

        sheet1.Calculate();

        Assert.IsTrue(sheet1.Cells["B2"].Value is ExcelErrorValue);
        Assert.IsTrue(sheet1.Cells["C3"].Value is bool);
        Assert.IsTrue((bool)sheet1.Cells["C3"].Value);
    }

    [TestMethod]
    public void I690InsertRow()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
        _ = pck.Workbook.Names.Add("ColumnRange", new ExcelRangeBase(sheet1, "Sheet1!$C:$C"));

        string? rangeAddress = pck.Workbook.Names["ColumnRange"].Address;

        sheet1.InsertRow(1, 1);

        Assert.AreEqual(pck.Workbook.Names["ColumnRange"].Address, rangeAddress);
    }

    [TestMethod]
    public void I690InsertColumn()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
        _ = pck.Workbook.Names.Add("RowRange", new ExcelRangeBase(sheet1, "Sheet1!$7:$7"));

        string? rangeAddress = pck.Workbook.Names["RowRange"].Address;

        sheet1.InsertColumn(1, 1);

        Assert.AreEqual(pck.Workbook.Names["RowRange"].Address, rangeAddress);
    }

    [TestMethod]
    public void I690DeleteRow()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
        _ = pck.Workbook.Names.Add("ColumnRange", new ExcelRangeBase(sheet1, $"Sheet1!$C2:$C{ExcelPackage.MaxRows}"));

        sheet1.DeleteRow(1, 1);

        Assert.AreEqual("Sheet1!$C:$C", pck.Workbook.Names["ColumnRange"].Address);
    }

    [TestMethod]
    public void I690DeleteColumn()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
        _ = pck.Workbook.Names.Add("RowRange", new ExcelRangeBase(sheet1, $"Sheet1!$B$7:$XFD$7"));

        sheet1.DeleteColumn(1, 1);

        Assert.AreEqual("Sheet1!$7:$7", pck.Workbook.Names["RowRange"].Address);
    }

    [TestMethod]
    public void s330()
    {
        using ExcelPackage? p = OpenTemplatePackage("s330.xlsm");
        p.Workbook.VbaProject.Signature.Certificate = null;
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i694()
    {
        using ExcelPackage? p = OpenPackage("i694.xlsx", true);
        ExcelWorkbook? wb = p.Workbook;
        ExcelWorksheet? ws = wb.Worksheets.Add("test");

        for (int colNum = 1; colNum <= 5; colNum++)
        {
            ws.Cells[1, colNum].Value = $"Column_{colNum}";
            ws.Column(colNum).OutlineLevel = 1;
            ws.Column(colNum).Collapsed = true;
            ws.Column(colNum).Hidden = true;
        }

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i701()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"i701.xlsx");
        ExcelTable partsTable = p.Workbook.Worksheets.SelectMany(x => x.Tables).SingleOrDefault(x => x.Name == "TblComponents");

        if (!partsTable.Columns.Any(x => x.Name == "TEST"))
        {
            _ = partsTable.Columns.Add();
        }

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void AddExternalWorkbookNoUpdate()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"extref_relative.xlsx");
        p.Workbook.ExternalLinks[0].As.ExternalWorkbook.IsPathRelative = true;
        _ = p.Workbook.ExternalLinks[0].As.ExternalWorkbook.Load();
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i715()
    {
        using (ExcelPackage? p = OpenTemplatePackage(@"i715.xlsx"))
        {
            ExcelTable? t = p.Workbook.Worksheets["Data"].Tables[0];
            int columnPosition = t.Columns.First(c => c.Name == "Extra").Position;
            _ = t.Columns.Delete(columnPosition, 1);
            SaveAndCleanup(p);
        }

        using (ExcelPackage? p = OpenTemplatePackage(@"i715-2.xlsx"))
        {
            ExcelTable? t = p.Workbook.Worksheets["Data"].Tables[0];
            int columnPosition = t.Columns.First(c => c.Name == "Extra").Position;
            _ = t.Columns.Delete(columnPosition, 1);

            // delete one more column
            columnPosition = t.Columns.First(c => c.Name == "Col X").Position;
            _ = t.Columns.Delete(columnPosition, 1);
            SaveAndCleanup(p);
        }

        using (ExcelPackage? p = OpenTemplatePackage(@"i715-3.xlsx"))
        {
            ExcelTable? t = p.Workbook.Worksheets["Bookings This Year"].Tables["JobsThisYear"];

            // if I delete even one column it produces corrupted xlsx
            int columnPosition = t.Columns.First(c => c.Name == "Total Time").Position;
            _ = t.Columns.Delete(columnPosition, 1);

            SaveAndCleanup(p);
        }
    }

    [TestMethod]
    public void i707()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"i707.xlsx");
        SaveAndCleanup(p);
    }

    public static void i720()
    {
        //_ = new MemoryStream();
        using ExcelPackage? p = OpenPackage("i720.xlsx");
        p.Settings.TextSettings.DefaultTextMeasurer = new OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics.DefaultTextMeasurer();
        ExcelWorksheet? worksheet = p.Workbook.Worksheets.Add("Sheet1");
        worksheet.Cells["A1"].Value = "Test";
        worksheet.Cells["a:xfd"].AutoFitColumns();
    }

    [TestMethod]
    public void i729()
    {
        string? id = "2";
        using ExcelPackage? p = OpenTemplatePackage($"i729-{id}.xlsm");
        string? path = $"c:\\temp\\vba-issue\\{id}\\";

        if (Directory.Exists(path))
        {
            Directory.Delete(path, true);
        }

        _ = Directory.CreateDirectory(path);

        StringWriter? sb = new StringWriter();
        WriteStorage(p.Workbook.VbaProject.Document.Storage, sb, path, "");

        SaveWorkbook("i729.xlsm", p);
    }

    [TestMethod]
    public void i735()
    {
        using ExcelPackage? p = OpenPackage("i735.xlsx", true);
        ExcelWorkbook? wb = p.Workbook;
        ExcelWorksheet? ws = wb.Worksheets.Add("test");
        ExcelRange? tableRange = ws.Cells["A1:C5"];
        tableRange.Style.Border.Top.Style = ExcelBorderStyle.Dashed;
        tableRange.Style.Border.Left.Style = ExcelBorderStyle.Dashed;
        tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
        tableRange.Style.Border.Right.Style = ExcelBorderStyle.Dashed;

        //I want row 2 to be a dotted edge. But only the right edge changes.
        ws.Cells["A2:C2"].Style.Border.BorderAround(ExcelBorderStyle.DashDotDot); //this doesn't work.
        SaveAndCleanup(p);
    }

    private static void WriteStorage(CompoundDocument.StoragePart storage, StringWriter sb, string path, string dir)
    {
        foreach (string? key in storage.SubStorage.Keys)
        {
            _ = Directory.CreateDirectory($"{path}{dir}{key}\\");
            WriteStorage(storage.SubStorage[key], sb, path, dir + $"{key}\\");
        }

        foreach (string? key in storage.DataStreams.Keys)
        {
            sb.WriteLine($"{path}{dir}\\" + key);
            File.WriteAllBytes($"{path}{dir}\\" + GetFileName(key) + ".bin", storage.DataStreams[key]);
        }
    }

    private static string GetFileName(string key)
    {
        StringBuilder? sb = new StringBuilder();
        char[]? ic = Path.GetInvalidFileNameChars();

        foreach (char c in key)
        {
            if (ic.Contains(c))
            {
                _ = sb.Append($"0x{(int)c}");
            }
            else
            {
                _ = sb.Append(c);
            }
        }

        return sb.ToString();
    }

    public class BorderInfo
    {
        public string RangeType { get; set; }
    }

    [TestMethod]
    public void i740()
    {
        using ExcelPackage? package = OpenPackage($"i738.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("sheet1");
        ExcelRange? tableRange = sheet.Cells["A1:C5"];

        List<BorderInfo> brders = new List<BorderInfo>()
        {
            new BorderInfo() { RangeType = "border-all" }, new BorderInfo() { RangeType = "border-outside" }, new BorderInfo() { RangeType = "border-none" }
        };

        foreach (BorderInfo? border in brders)
        {
            if (border.RangeType.Equals("border-all"))
            {
                tableRange.Style.Border.Top.Style = ExcelBorderStyle.Dashed;
                tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                tableRange.Style.Border.Left.Style = ExcelBorderStyle.Dashed;
                tableRange.Style.Border.Right.Style = ExcelBorderStyle.Dashed;
            }
            else if (border.RangeType.Equals("border-outside"))
            {
                tableRange.Style.Border.BorderAround(ExcelBorderStyle.Dashed);
            }
            else if (border.RangeType.Equals("border-none"))
            {
                tableRange["B2"].Style.Border.Top.Style = ExcelBorderStyle.None;
                tableRange["B2"].Style.Border.Bottom.Style = ExcelBorderStyle.None;
                tableRange["B2"].Style.Border.Left.Style = ExcelBorderStyle.None;
                tableRange["B2"].Style.Border.Right.Style = ExcelBorderStyle.None;

                tableRange["B1"].Style.Border.Bottom.Style = ExcelBorderStyle.None;
                tableRange["A2"].Style.Border.Right.Style = ExcelBorderStyle.None;
                tableRange["C2"].Style.Border.Left.Style = ExcelBorderStyle.None;
                tableRange["B3"].Style.Border.Top.Style = ExcelBorderStyle.None;
            }
        }

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void I742()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"i742.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I743()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"i743.xlsx");
        int[] rows = { 1, 3, 5, 7, 9, 12, 14, 16, 18, 20 };

        foreach (int row in rows)
        {
            ExcelRange? cell = p.Workbook.Worksheets[0].Cells[row, 3];
            string? cellColor = "";

            if (!string.IsNullOrEmpty(cell.Style.Font.Color.Theme.ToString()))
            {
                string? theme = cell.Style.Font.Color.Theme.ToString();

                if (theme == "Text1")
                {
                    theme = "Dark1";
                }
                else if (theme == "Text2")
                {
                    theme = "Dark2";
                }
                else if (theme == "Background1")
                {
                    theme = "Light1";
                }
                else if (theme == "Background2")
                {
                    theme = "Light2";
                }

                ExcelDrawingThemeColorManager? colorManager = (ExcelDrawingThemeColorManager)p.Workbook.ThemeManager.CurrentTheme.ColorScheme.GetType()
                                                                                              .GetProperty(theme)
                                                                                              .GetValue(p.Workbook.ThemeManager.CurrentTheme.ColorScheme);

                switch (colorManager.ColorType)
                {
                    case eDrawingColorType.None:
                        break;

                    case eDrawingColorType.RgbPercentage:
                        break;

                    case eDrawingColorType.Rgb:
                        cellColor = '#' + colorManager.RgbColor.Color.Name.ToUpper();

                        break;

                    case eDrawingColorType.Hsl:
                        break;

                    case eDrawingColorType.System:
                        cellColor = '#' + colorManager.SystemColor.LastColor.Name.ToUpper();

                        break;

                    case eDrawingColorType.Scheme:
                        break;

                    case eDrawingColorType.Preset:
                        break;

                    case eDrawingColorType.ChartStyleColor:
                        break;

                    default:
                        break;
                }
            }

            Debug.WriteLine(p.Workbook.Worksheets[0].Cells[row, 1].Text + $" address:{cell.Address}");
            Debug.WriteLine("Cell RGB: " + cell.Style.Font.Color.Rgb);
            Debug.WriteLine("Cell LookUp: " + cell.Style.Font.Color.LookupColor());
            Debug.WriteLine("Cell Theme: " + cellColor);

            if (cell.RichText != null)
            {
                foreach (ExcelRichText? richText in cell.RichText)
                {
                    Debug.WriteLine(richText.Text + " " + richText.Color.Name.ToUpper());
                }
            }
        }

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s404()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"s404.xlsm");
        p.Workbook.VbaProject.Protection.SetPassword(null);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i751()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"i751-Normal.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ws.DeleteColumn(3);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i752()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"i752.xlsx");

        for (int row = 1; row <= p.Workbook.Worksheets[0].Dimension.End.Row; row++)
        {
            if (p.Workbook.Worksheets[0].Cells[row, 1].Text == "")
            {
                continue;
            }

            ExcelRange? cell = p.Workbook.Worksheets[0].Cells[row, 3];

            Debug.WriteLine($"\n--> {p.Workbook.Worksheets[0].Cells[row, 1].Text}");
            Debug.Write($"Cell Font: [{cell.Style.Font.Name}], Size: [{cell.Style.Font.Size}]");

            if (cell.Style.Font.Bold)
            {
                Console.Write(", Bold");
            }

            Debug.WriteLine("");

            foreach (ExcelRichText? richText in cell.RichText)
            {
                Debug.Write($"RichText {richText.Text} Font: [{richText.FontName}], Size: [{richText.Size}]");

                if (richText.Bold)
                {
                    Console.Write(", Bold");
                }

                Debug.WriteLine("");
            }
        }
    }

    [TestMethod]
    public void s407()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"s407.xlsx");
        ExcelWorksheet? sheet1 = p.Workbook.Worksheets.First();
        _ = p.Workbook.Worksheets.Add("copy", sheet1);
        p.Save();
    }

    [TestMethod]
    public void i761()
    {
        using ExcelPackage? package = OpenPackage("i761.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Page1");
        sheet.Cells["A1"].Value = 0;
        sheet.Cells["B1"].Value = 5;
        sheet.Cells["A2"].Value = 0;
        sheet.Cells["B2"].Value = 1;
        sheet.Cells["A3"].Value = 1;
        sheet.Cells["B3"].Value = 1;
        sheet.Cells["A4"].Value = 1;
        sheet.Cells["B4"].Value = 4;
        sheet.Cells["F1"].Value = 1;
        sheet.Cells["G1"].CreateArrayFormula($"MIN(IF(A1:A4=F1,B1:B4))");
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void i762()
    {
        using ExcelPackage? p = OpenTemplatePackage(@"i762.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ws.Cells["C5"].Value = 15;
        ExcelPieChart? pc = ws.Drawings[0].As.Chart.PieChart;
        pc.Series[0].CreateCache();

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i763()
    {
        using ExcelPackage? package = OpenPackage("i761.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Page1");
        sheet.Cells["A1"].Value = 0;
        sheet.Cells["A2"].Value = 5;
        sheet.Cells["A3"].Value = 0;
        sheet.Cells["A4"].Value = 1;
        sheet.Cells["A5"].Value = 1;
        sheet.Cells["A6"].Value = 1;
        sheet.Cells["A7"].Value = 1;
        sheet.Cells["A8"].Value = 4;
        sheet.Cells["A9"].Value = 1;

        foreach (ExcelRangeBase? c in sheet.Cells["A1:A3,A5,A6,A7"])
        {
            Console.WriteLine($"{c.Address}");
        }
    }

    [TestMethod]
    public void ExcelRangeBase_Counts_Periods_Twice_763_1()
    {
        using ExcelPackage? p = new ExcelPackage();
        _ = p.Workbook.Worksheets.Add("first");

        ExcelWorksheet? sheet = p.Workbook.Worksheets.First();

        sheet.Cells["A1"].Value = 1;
        sheet.Cells["A2"].Value = 1;
        sheet.Cells["A3"].Value = 1;
        sheet.Cells["A4"].Value = 1;
        sheet.Cells["A5"].Value = 1;
        sheet.Cells["A6"].Value = 1;
        sheet.Cells["A7"].Value = 1;
        sheet.Cells["A8"].Value = 1;
        sheet.Cells["A9"].Value = 1;
        sheet.Cells["A10"].Value = 1;
        sheet.Cells["A11"].Value = 1;
        sheet.Cells["A12"].Formula = "SUM(A1:A3,A5,A6,A7,A8,A10,A9,A11)";

        int counterFirstIteration = 0;
        int counterSecondIteration = 0;

        int CounterSingleAdress = 0;
        int CounterMultipleRanges = 0;
        int CounterRangesFirst = 0;
        int CounterRangesLast = 0;
        int counterNoRanges = 0;
        int counterOneRange = 0;
        int counterMixed = 0;
        string? cellsFirstIteration = string.Empty;
        string? cellsSecondIteration = string.Empty;
        string? cellsSingleAdress = string.Empty;
        string? cellsMultipleRanges = string.Empty;
        string? cellsRangesFirst = string.Empty;
        string? cellsRangesLast = string.Empty;
        string? cellsNoRanges = string.Empty;
        string? cellsOneRange = string.Empty;
        string? cellsMixed = string.Empty;

        //------------------
        string? cellsInRange = string.Empty;

        ExcelRange? rangeWithPeriod = sheet.Cells["A1:A3,A5,A6,A7,A8,A10,A9,A11"];

        foreach (ExcelRangeBase? cell in rangeWithPeriod)
        {
            cellsInRange = $"{cellsInRange};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A5;A6;A7;A8;A10;A9;A11", cellsInRange);

        //------------------

        ExcelRange? range = sheet.Cells["A1:A3,A5,A6,A7,A8,A10,A9,A11"];

        foreach (ExcelRangeBase? cell in range)
        {
            counterFirstIteration++;
            cellsFirstIteration = $"{cellsFirstIteration};{cell.Address}";
        }

        foreach (ExcelRangeBase? cell in range)
        {
            counterSecondIteration++;
            cellsSecondIteration = $"{cellsSecondIteration};{cell.Address}";
        }

        Assert.AreEqual(cellsFirstIteration, cellsSecondIteration);
        Assert.AreEqual(";A1;A2;A3;A5;A6;A7;A8;A10;A9;A11", cellsFirstIteration);

        Assert.AreEqual(counterFirstIteration, counterSecondIteration);
        Assert.AreEqual(10, counterFirstIteration);

        ExcelRange? rangeSingleAdress = sheet.Cells["A1"];

        foreach (ExcelRangeBase? cell in rangeSingleAdress)
        {
            CounterSingleAdress++;
            cellsSingleAdress = $"{cellsSingleAdress};{cell.Address}";
        }

        Assert.AreEqual(";A1", cellsSingleAdress);
        Assert.AreEqual(1, CounterSingleAdress);

        cellsSingleAdress = string.Empty;
        CounterSingleAdress = 0;

        foreach (ExcelRangeBase? cell in rangeSingleAdress)
        {
            CounterSingleAdress++;
            cellsSingleAdress = $"{cellsSingleAdress};{cell.Address}";
        }

        Assert.AreEqual(";A1", cellsSingleAdress);
        Assert.AreEqual(1, CounterSingleAdress);

        ExcelRange? rangeMultipleRanges = sheet.Cells["A1:A4,A5:A7,A8:A11"];

        foreach (ExcelRangeBase? cell in rangeMultipleRanges)
        {
            CounterMultipleRanges++;
            cellsMultipleRanges = $"{cellsMultipleRanges};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11", cellsMultipleRanges);
        Assert.AreEqual(11, CounterMultipleRanges);

        CounterMultipleRanges = 0;
        cellsMultipleRanges = string.Empty;

        foreach (ExcelRangeBase? cell in rangeMultipleRanges)
        {
            CounterMultipleRanges++;
            cellsMultipleRanges = $"{cellsMultipleRanges};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11", cellsMultipleRanges);
        Assert.AreEqual(11, CounterMultipleRanges);

        ExcelRange? rangeRangeFirst = sheet.Cells["A1:A4,A5,A6,A7"];

        foreach (ExcelRangeBase? cell in rangeRangeFirst)
        {
            CounterRangesFirst++;
            cellsRangesFirst = $"{cellsRangesFirst};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsRangesFirst);
        Assert.AreEqual(7, CounterRangesFirst);

        CounterRangesFirst = 0;
        cellsRangesFirst = string.Empty;

        foreach (ExcelRangeBase? cell in rangeRangeFirst)
        {
            CounterRangesFirst++;
            cellsRangesFirst = $"{cellsRangesFirst};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsRangesFirst);
        Assert.AreEqual(7, CounterRangesFirst);

        ExcelRange? rangeRangeLast = sheet.Cells["A1,A2,A3,A4:A7"];

        foreach (ExcelRangeBase? cell in rangeRangeLast)
        {
            CounterRangesLast++;
            cellsRangesLast = $"{cellsRangesLast};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsRangesLast);
        Assert.AreEqual(7, CounterRangesLast);

        CounterRangesLast = 0;
        cellsRangesLast = string.Empty;

        foreach (ExcelRangeBase? cell in rangeRangeLast)
        {
            CounterRangesLast++;
            cellsRangesLast = $"{cellsRangesLast};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsRangesLast);
        Assert.AreEqual(7, CounterRangesLast);

        ExcelRange? rangeOneRange = sheet.Cells["A1:A7"];

        foreach (ExcelRangeBase? cell in rangeOneRange)
        {
            counterOneRange++;
            cellsOneRange = $"{cellsOneRange};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsOneRange);
        Assert.AreEqual(7, counterOneRange);

        ExcelRange? rangeNoRange = sheet.Cells["A1,A2,A3,A4"];

        foreach (ExcelRangeBase? cell in rangeNoRange)
        {
            counterNoRanges++;
            cellsNoRanges = $"{cellsNoRanges};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4", cellsNoRanges);
        Assert.AreEqual(4, counterNoRanges);

        ExcelRange? rangeMixed = sheet.Cells["A1,A2,A3:A5,A6,A7"];

        foreach (ExcelRangeBase? cell in rangeMixed)
        {
            counterMixed++;
            cellsMixed = $"{cellsMixed};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsMixed);
        Assert.AreEqual(7, counterMixed);

        int counter = 0;
        string cells = string.Empty;

        foreach (ExcelRangeBase? cell in sheet.Cells)
        {
            counter++;
            cells = $"{cells};{cell.Address}";
        }

        Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12", cells);
        Assert.AreEqual(12, counter);
    }

    [TestMethod]
    public void I778()
    {
        using ExcelPackage? package = OpenTemplatePackage("i778.xlsx");
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void StyleKeepBoxes()
    {
        using ExcelPackage? p = OpenTemplatePackage("XfsStyles.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void BuildInStylesRegional()
    {
        using ExcelPackage? p = OpenPackage("BuildinStylesRegional.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1:A4"].FillNumber(1);
        ws.Cells["A1"].Style.Numberformat.Format = "#,##0 ;(#,##0)";
        ws.Cells["A2"].Style.Numberformat.Format = "#,##0 ;[Red](#,##0)";
        ws.Cells["A3"].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
        ws.Cells["A4"].Style.Numberformat.Format = "#,##0.00;[Red](#,##0.00)";
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void I809()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("DateSheet");

        ws.Cells["A1"].Value = "2022-11-25";

        ws.Cells["B1"].Value = "2022-11-25";
        ws.Cells["B2"].Value = "2022-11-30";
        ws.Cells["B3"].Value = "2022-11-25";

        ws.Cells["C1"].Formula = "=DAY(B1)";
        ws.Cells["C2"].Formula = "=DAY(B2)";
        ws.Cells["C3"].Formula = "=DAY(B3)";

        ws.Cells["D1"].Formula = "SUMIF(B1:B3,A1,C1:C3)";
        p.Workbook.Calculate();

        Assert.AreEqual(50d, p.Workbook.Worksheets[0].Cells["D1"].Value);
    }

    [TestMethod]
    public void s425()
    {
        using ExcelPackage? p = OpenTemplatePackage("s425.xlsx");
        Assert.AreEqual(1, p.Workbook.Worksheets[0].PivotTables.Count);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s803()
    {
        using ExcelPackage? p = OpenPackage("s803.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1:E100"].FillNumber(1, 1);
        ws.Cells["A82"].Formula = "a1";
        ws.Cells["A82"].Formula = null;
        ws.Cells["A2:C100"].Style.Font.Bold = true;
        ws.InsertRow(81, 1);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s431()
    {
        using ExcelPackage? package = OpenTemplatePackage("template-not-working.xlsx");

        List<ExcelPicture>? pics =
            package.Workbook.Worksheets.SelectMany(p => p.Drawings).Where(p => p is ExcelPicture).Select(p => p as ExcelPicture).ToList();

        ExcelPicture? pic = pics.First(p => p.Name == "Image_ExistingInventoryImg");
        byte[]? image = File.ReadAllBytes("c:\\temp\\img1.png");
        _ = pic.Image.SetImage(image, ePictureType.Png);
        image = File.ReadAllBytes("c:\\temp\\img2.png");
        _ = pics[1].Image.SetImage(image, ePictureType.Png);

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void s430()
    {
        using ExcelPackage? package = OpenTemplatePackage("s430.xlsx");
        string? SheetName = "Error_Sheet";
        ExcelWorksheet? InSheet = package.Workbook.Worksheets[SheetName];

        using ExcelPackage? outPackage = OpenPackage("s430_out.xlsx", true);
        _ = outPackage.Workbook.Worksheets.Add(SheetName, InSheet);

        SaveAndCleanup(outPackage);
    }

    [TestMethod]
    public void s435()
    {
        using ExcelPackage? package = OpenTemplatePackage("s435.xlsx");
        ExcelWorksheet? worksheet = package.Workbook.Worksheets[0];

        int start = 2;
        worksheet.Cells["A" + start].Value = "Month";
        worksheet.Cells["B" + start].Value = "Serie1";
        worksheet.Cells["C" + start].Value = "Serie2";
        worksheet.Cells["D" + start].Value = "Serie3";
        start++;
        DataTable? randomData = Data();

        foreach (DataRow row in randomData.Rows)
        {
            worksheet.Cells["A" + start].Value = row[0];
            worksheet.Cells["B" + start].Value = row[1];
            worksheet.Cells["C" + start].Value = row[2];
            worksheet.Cells["D" + start].Value = row[3];
            start++;
        }

        int end = start - 1;
        ExcelChart metroChart = worksheet.Drawings.OfType<ExcelChart>().First();

        if (metroChart != null)
        {
            metroChart.YAxis.MinValue = 0.5;
            metroChart.YAxis.MaxValue = 4.5;

            Dictionary<string, string>? serieColors = new Dictionary<string, string>()
            {
                { "Serie1", "#CBB54C" }, { "Serie2", "#00A7CE" }, { "Serie3", "#950000" }
            };

            AddLineSeries(metroChart, $"B3:B{end}", $"A3:A{end}", "Serie1");
            AddLineSeries(metroChart, $"C3:C{end}", $"A3:A{end}", "Serie2");
            AddLineSeries(metroChart, $"D3:D{end}", $"A3:A{end}", "Serie3");

            foreach (ExcelLineChartSerie serie in metroChart.Series)
            {
                string? serieColor = serieColors[serie.Header];
                serie.Smooth = true;
                serie.Border.Fill.Color = ColorTranslator.FromHtml(serieColor);
                serie.Marker.Style = eMarkerStyle.None;
                serie.Border.Width = 2;
                serie.Border.Fill.Style = eFillStyle.SolidFill;
                serie.Border.LineStyle = eLineStyle.Solid;
                serie.Border.LineCap = eLineCap.Round;
                serie.Fill.Style = eFillStyle.SolidFill;
                serie.Fill.Color = ColorTranslator.FromHtml(serieColor);
            }
        }

        SaveAndCleanup(package);
    }

    private static void AddLineSeries(ExcelChart chart, string seriesAddress, string xSeriesAddress, string seriesName)
    {
        ExcelChartSerie? lineSeries = chart.Series.Add(seriesAddress, xSeriesAddress);
        lineSeries.Header = seriesName;
    }

    private static DataTable Data()
    {
        DataTable? toReturn = new DataTable();
        _ = toReturn.Columns.Add("Month");
        _ = toReturn.Columns.Add("Serie1", typeof(decimal));
        _ = toReturn.Columns.Add("Serie2", typeof(decimal));
        _ = toReturn.Columns.Add("Serie3", typeof(decimal));

        _ = toReturn.Rows.Add("01/2022", 1.4, 2.4, 3.4);
        _ = toReturn.Rows.Add("02/2022", 1.4, 2.4, 3.4);
        _ = toReturn.Rows.Add("03/2022", 1.4, 2.4, 3.4);
        _ = toReturn.Rows.Add("04/2022", 1.7, 2.7, 3.7);
        _ = toReturn.Rows.Add("05/2022", 1.7, 2.7, 3.7);
        _ = toReturn.Rows.Add("06/2022", 1.7, 2.7, 3.7);
        _ = toReturn.Rows.Add("07/2022", 1.9, 2.9, 3.9);
        _ = toReturn.Rows.Add("08/2022", 1.9, 2.9, 3.9);

        return toReturn;
    }

    [TestMethod]
    public void s437()
    {
        using ExcelPackage? package = OpenTemplatePackage("s437.xlsx");
        ExcelWorksheet? worksheet = package.Workbook.Worksheets[0];

        ExcelChart metroChart = worksheet.Drawings.OfType<ExcelChart>().First();

        if (metroChart != null)
        {
            metroChart.YAxis.MinValue = 0d;
            metroChart.YAxis.MajorUnit = 0.05d;

            foreach (ExcelChart? ct in metroChart.PlotArea.ChartTypes)
            {
                ///The "Series" being returned in this is only the bar series
                ///while the other two line series are not being returned.
                foreach (ExcelChartSerie? serie in ct.Series)
                {
                }
            }
        }
    }

    [TestMethod]
    public void Issue831()
    {
        using ExcelPackage package = OpenTemplatePackage("I831.xlsx");
        package.Workbook.Worksheets.Delete(0);
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void RichText()
    {
        using ExcelPackage package = OpenPackage("richtextclear.xlsx", true);
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.Red);

        _ = ws.Cells["A1"].RichText.Add("Test");
        ws.Cells["A1"].RichText.Clear();

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void extLst()
    {
        using ExcelPackage package = OpenTemplatePackage("extLstMany.xlsx");

        //package.Workbook.Worksheets.Delete(0);
        Assert.AreEqual(1, package.Workbook.Worksheets[0].DataValidations.Count);
        SaveAndCleanup(package);
    }

    //Should not generate a corrupt file when opened.
    [TestMethod]
    public void Issue842()
    {
        using ExcelPackage? package = OpenPackage("issue842_SHOULD_BE_OPENEABLE.xlsx", true);
        ExcelWorksheet ws = package.Workbook.Worksheets.Add("exampleSheet");

        ws.SetValue("C1", "0");

        ExcelAddress? address = new ExcelAddress(1, 1, 3, 1);
        IExcelDataValidationDecimal? testValidation = ws.DataValidations.AddDecimalValidation(address.Address);
        testValidation.Formula.Value = 10;
        testValidation.Formula2.Value = 15;

        ExcelRange range = ws.Cells[1, 1, 4, 1];
        ExcelTable table = ws.Tables.Add(range, "TestTable");
        table.StyleName = "None";

        _ = ws.Drawings.AddTableSlicer(table.Columns[0]);

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void s449()
    {
        using ExcelPackage? xlPackage = OpenTemplatePackage("s449.xlsx");
        using ExcelPackage? p2 = OpenPackage("s449-saved.xlsx", true);
        string? SheetName = "Error_Sheet";
        ExcelWorksheet? InSheet = xlPackage.Workbook.Worksheets[SheetName];
        _ = p2.Workbook.Worksheets.Add(SheetName, InSheet);
        SaveAndCleanup(p2);
    }

    [TestMethod]
    public void i848()
    {
        using ExcelPackage? p = OpenTemplatePackage("i848.xlsx");

        // We have 2 rows with formulas in C column.
        ExcelWorkbook? book = p.Workbook;
        ExcelWorksheet? eppWorksheet = book.Worksheets[0];

        // Insert row(-s) after first one. 
        // New row one-based index is 2.
        // Second row index is 3 now.
        eppWorksheet.InsertRow(2, 1);
        Assert.AreEqual("A3*B3", eppWorksheet.Cells["C3"].Formula);

        // Formula updated fine =A3*B3.
        Console.WriteLine(eppWorksheet.Cells["C3"].Formula);

        // Insert row(-s) after second one (it has index 3 now).
        eppWorksheet.InsertRow(4, 1);

        // Formula should not be updated, because row 3 is above.
        // But now its =A2*B2. Why?
        Assert.AreEqual("A3*B3", eppWorksheet.Cells["C3"].Formula);

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue848()
    {
        using ExcelPackage? p = OpenPackage("issue848-2.xlsx", true);
        ExcelWorksheet? worksheet = p.Workbook.Worksheets.Add("Sheet1");

        // We start with a single row that has a formula
        worksheet.Cells["A1"].Value = 1;
        worksheet.Cells["B1"].Value = 2;
        worksheet.Cells["C1"].Formula = "A1*B1";

        // Insert a row and copy the original row
        worksheet.InsertRow(2, 1);
        worksheet.Cells["A1:C1"].Copy(worksheet.Cells["A2:C2"]);
        Assert.AreEqual("A1*B1", worksheet.Cells["C1"].Formula);
        Assert.AreEqual("A2*B2", worksheet.Cells["C2"].Formula);

        // Insert another row and copy the original row
        worksheet.InsertRow(3, 1);
        worksheet.Cells["A1:C1"].Copy(worksheet.Cells["A3:C3"]);
        Assert.AreEqual("A1*B1", worksheet.Cells["C1"].Formula);
        Assert.AreEqual("A2*B2", worksheet.Cells["C2"].Formula);
        Assert.AreEqual("A3*B3", worksheet.Cells["C3"].Formula);

        // Delete the original row
        worksheet.DeleteRow(1);
        Assert.AreEqual("A1*B1", worksheet.Cells["C1"].Formula); // This still succeeds...
        Assert.AreEqual("A2*B2", worksheet.Cells["C2"].Formula);

        // Insert a row the end
        worksheet.InsertRow(4, 1); // This seems to trigger the issue

        // Next line fails, the formula in C1 is "A2*B2" (which is wrong)
        Assert.AreEqual("A1*B1", worksheet.Cells["C1"].Formula); //  ... Now this fails
        Assert.AreEqual("A2*B2", worksheet.Cells["C2"].Formula);
    }

    [TestMethod]
    public void Issue854()
    {
        //using (var p = OpenTemplatePackage("i854.xlsx"))
        using ExcelPackage? p = OpenTemplatePackage("i854.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["Component Failure Rates"];
        _ = ws.Tables[0].DeleteRow(0, 1);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue854_2_Insert()
    {
        using ExcelPackage? p = OpenTemplatePackage("i854-2.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["Components"];
        ExcelTable? table = ws.Tables[0];
        _ = table.Columns.Insert(1, 2);
        SaveWorkbook("i854-2-Insert.xlsx", p);
    }

    [TestMethod]
    public void Issue854_2_Delete()
    {
        using ExcelPackage? p = OpenTemplatePackage("i854-2.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets["Components"];
        ExcelTable? table = ws.Tables[0];
        _ = table.Columns.Delete(1, 2);
        SaveWorkbook("i854-2-Delete.xlsx", p);
    }

    [TestMethod]
    public void Issue854_3()
    {
        using ExcelPackage? p = OpenTemplatePackage("i854-3.xlsx");
        ExcelTable? table = p.Workbook.Worksheets.SelectMany(x => x.Tables).Single(x => x.Name == "TblComponents");

        List<object[]> tableData = new List<object[]>()
        {
            new[] { "C1", null, "Ceramic Capacitor", "Ceramic Capacitor FM" }, new[] { "C2", null, "Ceramic Capacitor", "Ceramic Capacitor FM" }
        };

        InsertTableRows(table, tableData.ToList(), 1, true);
        SaveAndCleanup(p);
    }

    static void InsertTableRows(ExcelTable table, List<object[]> data, int insertBeforeRow, bool removeOtherRows)
    {
        // Finding the row in the sheet
        int nextSheetRow = GetSheetRowIndex(table, insertBeforeRow);

        _ = table.InsertRow(insertBeforeRow, data.Count, true);

        // Applying the formulas to the new rows
        if (insertBeforeRow > 0)
        {
            nextSheetRow = GetSheetRowIndex(table, insertBeforeRow + data.Count);
        }

        // Filling the data -> not working
        for (int dataRowIx = 0; dataRowIx < data.Count; dataRowIx++)
        {
            object[] dataRow = data[dataRowIx];

            for (int colIx = 0; colIx < dataRow.Length; colIx++)
            {
                table.WorkSheet.Cells[nextSheetRow, colIx + 1].Value = dataRow[colIx];
            }

            nextSheetRow++;
        }

        // Deleting the other rows in the table if required
        if (removeOtherRows)
        {
            // Deleting the rows before the newly inserted ones
            if (insertBeforeRow > 0)
            {
                _ = table.DeleteRow(0, insertBeforeRow);
            }

            // NB: The new rows are now at the start of the sequence, so only the ones after
            // them should be removed
            int endRowsToRemove = table.Address.End.Row - table.Address.Start.Row - data.Count;

            if (endRowsToRemove > 0)
            {
                _ = table.DeleteRow(data.Count, endRowsToRemove);
            }
        }
    }

    static int GetSheetRowIndex(ExcelTable table, int tableRow) => table.Address.Start.Row + tableRow + (table.ShowHeader ? 1 : 0);

    [TestMethod]
    public void Issue852()
    {
        using ExcelPackage? p = OpenPackage("i852.xlsx", true);

        //Making the sheet
        ExcelWorksheet? sheet = p.Workbook.Worksheets.Add("mergedCellsTest");

        ExcelRange? cells = sheet.Cells["A1:D1"];

        cells.Merge = true;
        cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cells.Style.Fill.BackgroundColor.SetColor(Color.Orange);

        sheet.Cells["A2"].Value = "a";
        sheet.Cells["B2"].Value = "b";
        sheet.Cells["A3"].Value = "c";
        sheet.Cells["B3"].Value = "d";

        ExcelRange? square = sheet.Cells["C2:C3"];

        square.Merge = true;

        square.Style.Fill.PatternType = ExcelFillStyle.Solid;
        square.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

        ExcelRange? numberRange = sheet.Cells["A5:E5"];

        for (int i = 0; i < numberRange.Columns; i++)
        {
            numberRange.SetCellValue(0, i, i + 1);
        }

        ExcelRange? lowMerge = sheet.Cells["A6:E6"];

        lowMerge.Merge = true;

        lowMerge.Style.Fill.PatternType = ExcelFillStyle.Solid;
        lowMerge.Style.Fill.BackgroundColor.SetColor(Color.Blue);

        sheet.Cells["A2:C3"].Insert(eShiftTypeInsert.Right);

        Assert.IsTrue(sheet.Cells["A1:D1"].Merge);
        Assert.IsTrue(sheet.Cells["A6:E6"].Merge);
        Assert.IsFalse(sheet.Cells["E1:G1"].Merge);
        Assert.IsFalse(sheet.Cells["A6:H6"].Merge);

        Assert.AreEqual(Color.LightYellow.ToArgb().ToString("X"), sheet.Cells["F2:F3"].Style.Fill.BackgroundColor.Rgb);

        MemoryStream? stream = new MemoryStream();
        p.SaveAs(stream);

        ExcelPackage newPack = new ExcelPackage(stream);

        ExcelWorksheet? newSheet = newPack.Workbook.Worksheets[0];

        //Performing the test

        Assert.IsTrue(newSheet.Cells["A1:D1"].Merge);
        Assert.IsTrue(newSheet.Cells["A6:E6"].Merge);
        Assert.IsFalse(newSheet.Cells["E1:G1"].Merge);
        Assert.IsFalse(newSheet.Cells["A6:H6"].Merge);

        Assert.AreEqual(Color.LightYellow.ToArgb().ToString("X"), newSheet.Cells["F2:F3"].Style.Fill.BackgroundColor.Rgb);
    }

    [TestMethod]
    public void Issue858()
    {
        using ExcelPackage? p = OpenTemplatePackage("i858.xlsx");
        ExcelWorksheet? sheet = p.Workbook.Worksheets[0];

        for (int i = 1; i <= 4; i++)
        {
            sheet.Cells[1, i].IsRichText = true;
            sheet.Cells[1, i].Style.WrapText = true;
            sheet.Cells[1, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[1, i].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            _ = sheet.Cells[1, i].RichText.Add($"Hello world {i}");
        }

        _ = p.GetAsByteArray();
    }

    [TestMethod]
    public void Issue861()
    {
        ExcelPackage? ep = new ExcelPackage();
        ExcelWorksheet? ws = ep.Workbook.Worksheets.Add("Test");

        for (int row = 1; row < 10; ++row)
        {
            for (int col = 1; col < 10; ++col)
            {
                ws.Cells[row, col].Value = $"{row}:{col}";
            }
        }

        ExcelColumn? wsCol = ws.Column(3);
        wsCol.Style.Border.Left.Style = wsCol.Style.Border.Right.Style = ExcelBorderStyle.Thick;
        wsCol.Style.Fill.SetBackground(Color.Black);

        ws.Row(3).Style.Fill.SetBackground(Color.Aqua);

        Assert.AreNotEqual(ws.Row(3).Style.Border.Left.Style, wsCol.Style.Border.Left.Style);
        Assert.AreNotEqual(ws.Row(3).Style.Border.Right.Style, wsCol.Style.Border.Right.Style);
    }

    [TestMethod]
    public void i863()
    {
        using ExcelPackage? p = OpenTemplatePackage("i863.xlsx");

        // Removed insertion of PHI data, just re-saving the template for sample purposes

        // Workaround - Issue with "Inputs" tab - Validation of T60:T64 failed: Formula2 must be set if operator is 'between' or 'notBetween' when cells are not using between or notBetween
        ExcelWorksheet? otherInputTab = p.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals("Inputs"));

        if (otherInputTab != null)
        {
            otherInputTab.DataValidations.InternalValidationEnabled = false;
        }

        // Saving
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void Issue864()
    {
        using ExcelPackage? p = OpenPackage("i864.xlsx", true);
        ExcelWorksheet? worksheet = p.Workbook.Worksheets.Add("Test");

        worksheet.Cells[1, 1].Value = 1100;
        worksheet.Cells[1, 2].Value = 2200;
        worksheet.Cells[1, 3].Value = 3300;
        worksheet.Cells[1, 4].Value = 4400;
        worksheet.Cells[2, 1].Value = 2000;
        worksheet.Cells[2, 2].Value = 1254;
        worksheet.Cells[2, 3].Value = 5423;
        worksheet.Cells[2, 4].Value = 1400;
        worksheet.Cells[3, 1].Value = 2343;
        worksheet.Cells[3, 2].Value = 2355;
        worksheet.Cells[3, 3].Value = 2121;
        worksheet.Cells[3, 4].Value = 1231;
        worksheet.Cells[4, 1].Value = 954;
        worksheet.Cells[4, 2].Value = 4323;
        worksheet.Cells[4, 3].Value = 1112;
        worksheet.Cells[4, 4].Value = 2211;

        ExcelSparklineGroup? sparklineGroups = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells[1, 6, 4, 6], worksheet.Cells[1, 1, 4, 4]);

        sparklineGroups.MinAxisType = eSparklineAxisMinMax.Custom;
        sparklineGroups.ManualMin = 0;
        sparklineGroups.MaxAxisType = eSparklineAxisMinMax.Group;
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s461()
    {
        using ExcelPackage? p = OpenTemplatePackage("s461.xlsx");
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s463()
    {
        using ExcelPackage? p = OpenTemplatePackage("SRK2016.xlsx");

        foreach (ExcelWorksheet? ws in p.Workbook.Worksheets)
        {
            if (ws.Names.ContainsKey("_xlnm.Print_Area") && ws.Names.ContainsKey("Print_Area"))
            {
                ws.Names.Remove("Print_Area");
            }
        }

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void i871()
    {
        using ExcelPackage? p = OpenTemplatePackage("i871.xlsx");
        ExcelTable? table = p.Workbook.Worksheets.SelectMany(x => x.Tables).Single(x => x.Name == "TblComponentTypes");
        _ = table.AddRow(2);
        table.WorkSheet.Cells[8, 1].Value = "TST";
        table.WorkSheet.Cells[9, 1].Value = "TST 2";

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void s466()
    {
        using ExcelPackage? package = OpenTemplatePackage("s466.xlsx");
        ExcelTable? table = package.Workbook.Worksheets.SelectMany(x => x.Tables).Single(x => x.Name == "TblEffects");
        _ = table.AddRow(2);
        table.WorkSheet.Cells[8, 1].Value = "TST";
        table.WorkSheet.Cells[8, 2].Value = "Safe";
        table.WorkSheet.Cells[9, 1].Value = "TST 2";
        table.WorkSheet.Cells[9, 2].Value = "Dangerous";

        _ = table.InsertRow(0, 5);
        _ = table.DeleteRow(0, 5);
        SaveAndCleanup(package);
    }
}