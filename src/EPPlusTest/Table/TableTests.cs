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
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace EPPlusTest.Table;

[TestClass]
public class TableTests : TestBase
{
    static ExcelPackage _pck;
    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("Table.xlsx", true);
    }
    [ClassCleanup]
    public static void Cleanup()
    {
        SaveAndCleanup(_pck);
    }

    [TestMethod]
    public void TableWithSubtotalsParensInColumnName()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableSubtotParensColumnName");
        ws.Cells["B2"].Value = "Header 1";
        ws.Cells["C2"].Value = "Header (2)";
        ws.Cells["B3"].Value = 1;
        ws.Cells["B4"].Value = 2;
        ws.Cells["C3"].Value = 3;
        ws.Cells["C4"].Value = 4;
        ExcelTable? table = ws.Tables.Add(ws.Cells["B2:C4"], "TestTableParamHeader");
        table.ShowTotal = true;
        table.ShowHeader = true;
        table.Columns[0].TotalsRowFunction = RowFunctions.Sum;
        table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
        ws.Cells["B5"].Calculate();
        Assert.AreEqual(3.0, ws.Cells["B5"].Value);
        ws.Cells["C5"].Calculate();
        Assert.AreEqual(7.0, ws.Cells["C5"].Value);
    }
    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void TestTableNameCanNotStartsWithNumber()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("Table");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1"], "5TestTable");
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void TestTableNameCanNotContainWhiteSpaces()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("TableNoWhiteSpace");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1"], "Test Table");
    }

    [TestMethod]
    public void TestTableNameCanStartsWithBackSlash()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NameStartWithBackSlash");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1"], "\\TestTable");
    }

    [TestMethod]
    public void TestTableNameCanStartsWithUnderscore()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NameStartWithUnderscore");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1"], "_TestTable");
    }
    [TestMethod]
    public void TableTotalsRowFunctionEscapesSpecialCharactersInColumnName()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TotalsFormulaTest");
        ws.Cells["A1"].Value = "Col1";
        ws.Cells["B1"].Value = "[#'Col2']";
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:B2"], "TableFormulaTest");
        tbl.ShowTotal = true;
        tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
        Assert.AreEqual("SUBTOTAL(109,TableFormulaTest['['#''Col2''']])", ws.Cells["B3"].Formula);
    }
    [TestMethod]
    public void ValidateEncodingForTableColumnNames()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValidateTblColumnNames");
        ws.Cells["A1"].Value = "Col1>";
        ws.Cells["B1"].Value = "Col1&gt;";
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:C2"], "TableValColNames");
        Assert.AreEqual("Col1>", tbl.Columns[0].Name);
        Assert.AreEqual("Col1&gt;", tbl.Columns[1].Name);
        Assert.AreEqual("Column3", tbl.Columns[2].Name);
    }
    [TestMethod]
    public void TableTest()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Table");
        ws.Cells["B1"].Value = 123;
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["B1:P12"], "TestTable");
        tbl.TableStyle = TableStyles.Custom;

        tbl.ShowFirstColumn = true;
        tbl.ShowTotal = true;
        tbl.ShowHeader = true;
        tbl.ShowLastColumn = true;
        tbl.ShowFilter = false;
        Assert.AreEqual(tbl.ShowFilter, false);
        ws.Cells["K2"].Value = 5;
        ws.Cells["J3"].Value = 4;

        tbl.Columns[8].TotalsRowFunction = RowFunctions.Sum;
        tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])", tbl.Columns[9].Name);
        tbl.Columns[14].CalculatedColumnFormula = "TestTable[[#This Row],[123]]+TestTable[[#This Row],[Column2]]";
        ws.Cells["B2"].Value = 1;
        ws.Cells["B3"].Value = 2;
        ws.Cells["B4"].Value = 3;
        ws.Cells["B5"].Value = 4;
        ws.Cells["B6"].Value = 5;
        ws.Cells["B7"].Value = 6;
        ws.Cells["B8"].Value = 7;
        ws.Cells["B9"].Value = 8;
        ws.Cells["B10"].Value = 9;
        ws.Cells["B11"].Value = 10;
        ws.Cells["B12"].Value = 11;
        ws.Cells["O5"].Value = 11;
        ws.Cells["C7"].Value = "Table test";
        ws.Cells["C8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells["C8"].Style.Fill.BackgroundColor.SetColor(Color.Red);

        tbl = ws.Tables.Add(ws.Cells["a12:a13"], "");

        tbl = ws.Tables.Add(ws.Cells["C16:Y35"], "");
        tbl.TableStyle = TableStyles.Medium14;
        tbl.ShowFirstColumn = true;
        tbl.ShowLastColumn = true;
        tbl.ShowColumnStripes = true;
        Assert.AreEqual(tbl.ShowFilter, true);
        tbl.Columns[2].Name = "Test Column Name";

        ws.Cells["G50"].Value = "Timespan";
        ws.Cells["G51"].Value = new DateTime(new TimeSpan(1, 1, 10).Ticks); //new DateTime(1899, 12, 30, 1, 1, 10);
        ws.Cells["G52"].Value = new DateTime(1899, 12, 30, 2, 3, 10);
        ws.Cells["G53"].Value = new DateTime(1899, 12, 30, 3, 4, 10);
        ws.Cells["G54"].Value = new DateTime(1899, 12, 30, 4, 5, 10);

        ws.Cells["G51:G55"].Style.Numberformat.Format = "HH:MM:SS";
        tbl = ws.Tables.Add(ws.Cells["G50:G54"], "");
        tbl.ShowTotal = true;
        tbl.ShowFilter = false;
        tbl.Columns[0].TotalsRowFunction = RowFunctions.Sum;
    }

    [TestMethod]
    public void TableDeleteTest()
    {
        using ExcelPackage? p = OpenPackage("TableDeleteTest.xlsx", true);
        ExcelWorkbook? wb = p.Workbook;
        ExcelWorksheet[]? sheets = new[]
        {
            wb.Worksheets.Add("WorkSheet A"),
            wb.Worksheets.Add("WorkSheet B")
        };
        for (int i = 1; i <= 4; i++)
        {
            ExcelRange? cell = sheets[0].Cells[1, i];
            cell.Value = cell.Address + "_";
            cell = sheets[1].Cells[1, i];
            cell.Value = cell.Address + "_";
        }

        for (int i = 6; i <= 11; i++)
        {
            ExcelRange? cell = sheets[0].Cells[3, i];
            cell.Value = cell.Address + "_";
            cell = sheets[1].Cells[3, i];
            cell.Value = cell.Address + "_";
        }
        ExcelTable[]? tables = new[]
        {
            sheets[1].Tables.Add(sheets[1].Cells["A1:D73"], "TableDeletea"),
            sheets[0].Tables.Add(sheets[0].Cells["A1:D73"], "TableDelete2"),
            sheets[1].Tables.Add(sheets[1].Cells["F3:K10"], "TableDeleteb"),
            sheets[0].Tables.Add(sheets[0].Cells["F3:K10"], "TableDelete3"),
        };
        Assert.AreEqual(5, wb._nextTableID);
        Assert.AreEqual(1, tables[0].Id);
        Assert.AreEqual(2, tables[1].Id);
        try
        {
            sheets[0].Tables.Delete("TableDeletea");
            Assert.Fail("ArgumentException should have been thrown.");
        }
        catch (ArgumentOutOfRangeException) { }
        sheets[1].Tables.Delete("TableDeletea");
        Assert.AreEqual(1, tables[1].Id);
        Assert.AreEqual(2, tables[2].Id);

        try
        {
            sheets[1].Tables.Delete(4);
            Assert.Fail("ArgumentException should have been thrown.");
        }
        catch (ArgumentOutOfRangeException) { }
        ExcelRange? range = sheets[0].Cells[sheets[0].Tables[1].Address.Address];
        sheets[0].Tables.Delete(1, true);
        foreach (ExcelRangeBase? cell in range)
        {
            Assert.IsNull(cell.Value);
        }
        SaveAndCleanup(p);
    }
    [TestMethod]
    public void DeleteTablesFromTemplate()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Tablews1");
        ws.Tables.Add(new ExcelAddressBase("A1:C3"), "Table1");
        ws.Tables.Add(new ExcelAddressBase("D1:G7"), "Table2");

        Assert.AreEqual(2, ws.Tables.Count);
        p.Save();

        using ExcelPackage? p2 = new ExcelPackage(p.Stream);
        ws = p2.Workbook.Worksheets[0];
        Assert.AreEqual(2, ws.Tables.Count);
        ws.Tables.Delete(0);
        ws.Tables.Delete("Table2");

        Assert.AreEqual(0, ws.Tables.Count);
        p2.Save();
        using ExcelPackage? p3 = new ExcelPackage(p2.Stream);
        Assert.AreEqual(0, p3.Workbook.Worksheets[0].Tables.Count);
    }
    [TestMethod]
    public void ValidateTableSaveLoad()
    {
        using ExcelPackage? p1 = OpenPackage("table.xlsx");
        ExcelWorksheet? sheet = p1.Workbook.Worksheets.Add("Tables");

        // headers
        sheet.Cells["A1"].Value = "Month";
        sheet.Cells["B1"].Value = "Sales";
        sheet.Cells["C1"].Value = "VAT";
        sheet.Cells["D1"].Value = "Total";

        Random? rnd = new Random();
        for (int row = 2; row < 12; row++)
        {
            sheet.Cells[row, 1].Value = new DateTimeFormatInfo().GetMonthName(row);
            sheet.Cells[row, 2].Value = rnd.Next(10000, 100000);
            sheet.Cells[row, 3].Formula = $"B{row} * 0.25";
            sheet.Cells[row, 4].Formula = $"B{row} + C{row}";
        }
        sheet.Cells["B2:D13"].Style.Numberformat.Format = "€#,##0.00";

        ExcelRange? range = sheet.Cells["A1:D11"];

        // create the table
        ExcelTable? table = sheet.Tables.Add(range, "myTable");
        // configure the table
        table.ShowHeader = true;
        table.ShowFirstColumn = true;
        table.TableStyle = TableStyles.Dark2;
        // add a totals row under the data
        table.ShowTotal = true;
        table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
        table.Columns[2].TotalsRowFunction = RowFunctions.Sum;
        table.Columns[3].TotalsRowFunction = RowFunctions.Sum;

        // Calculate all the formulas including the totals row.
        // This will give input to the AutofitColumns call
        range.Calculate();
        range.AutoFitColumns();

        p1.Save();
        using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
        sheet = p2.Workbook.Worksheets["Tables"];
        // get a table by its name and change properties
        ExcelTable? myTable = sheet.Tables["myTable"];
        myTable.TableStyle = TableStyles.Medium8;
        myTable.ShowFirstColumn = false;
        myTable.ShowLastColumn = true;
        Assert.AreEqual(TableStyles.Medium8, myTable.TableStyle);
        SaveWorkbook("Table2.xlsx", p2);
        using ExcelPackage? p3 = new ExcelPackage(p2.Stream);
        sheet = p3.Workbook.Worksheets["Tables"];
        // get a table by its name and change properties
        sheet.Tables.Delete("myTable");

        SaveWorkbook("Table3.xlsx", p3);
    }
    [TestMethod]
    public void AddRowShouldAdjustSubtotals()
    {
        using ExcelPackage? package = OpenPackage("TableAdjustSubtotals.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Tables");

        // headers
        sheet.Cells["A1"].Value = "Month";
        sheet.Cells["B1"].Value = "Sales";
        sheet.Cells["C1"].Value = "VAT";
        sheet.Cells["D1"].Value = "Total";

        Random? rnd = new Random();
        for (int row = 2; row < 12; row++)
        {
            sheet.Cells[row, 1].Value = new DateTimeFormatInfo().GetMonthName(row);
            sheet.Cells[row, 2].Value = rnd.Next(10000, 100000);
            sheet.Cells[row, 3].Formula = $"B{row} * 0.25";
            sheet.Cells[row, 4].Formula = $"B{row} + C{row}";
        }
        sheet.Cells["B2:D13"].Style.Numberformat.Format = "€#,##0.00";

        ExcelRange? range = sheet.Cells["A1:D11"];

        // create the table
        ExcelTable? table = sheet.Tables.Add(range, "myTable");
        // configure the table
        table.ShowHeader = true;
        table.ShowFirstColumn = true;
        table.ShowFilter = false;
        table.TableStyle = TableStyles.Dark2;
        // add a totals row under the data
        table.ShowTotal = true;
        table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
        table.Columns[2].TotalsRowFunction = RowFunctions.Sum;
        table.Columns[3].TotalsRowFunction = RowFunctions.Sum;

        // insert rows
        ExcelRangeBase? rowRange = table.AddRow();
        int newRowIx = rowRange.Start.Row;
        sheet.Cells[newRowIx, 1].Value = new DateTimeFormatInfo().GetMonthName(newRowIx);
        sheet.Cells[newRowIx, 2].Value = rnd.Next(10000, 100000);
        sheet.Cells[newRowIx, 3].Formula = $"B{newRowIx} * 0.25";
        sheet.Cells[newRowIx, 4].Formula = $"B{newRowIx} + C{newRowIx}";

        // Calculate all the formulas including the totals row.
        sheet.Calculate();
        sheet.Cells.AutoFitColumns();

        SaveAndCleanup(package);
    }
    [TestMethod]
    public void ValidateCalculatedColumn()
    {
        using ExcelPackage? package = OpenPackage("TableCalculatedColumn.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Tables");

        // headers
        sheet.Cells["C1"].Value = "Month";
        sheet.Cells["D1"].Value = "Sales";
        sheet.Cells["E1"].Value = "VAT";
        sheet.Cells["F1"].Value = "Total";
        sheet.Cells["G1"].Value = "Formula";

        Random? rnd = new Random();
        for (int row = 2; row < 12; row++)
        {
            sheet.Cells[row, 3].Value = new DateTimeFormatInfo().GetMonthName(row);
            sheet.Cells[row, 4].Value = rnd.Next(10000, 100000);
            sheet.Cells[row, 5].Formula = $"D{row} * 0.25";
            sheet.Cells[row, 6].Formula = $"D{row} + E{row}";
        }
        sheet.Cells["D2:G13"].Style.Numberformat.Format = "€#,##0.00";

        ExcelRange? range = sheet.Cells["C1:G11"];

        // create the table
        ExcelTable? table = sheet.Tables.Add(range, "myTable");
        // configure the table
        table.ShowHeader = true;
        table.ShowTotal = true;

        string? formula = "mytable[[#this row],[Sales]]+mytable[[#this row],[VAT]]";
        table.Columns[4].CalculatedColumnFormula = formula;

        //Assert
        Assert.AreEqual(formula, table.Columns[4].CalculatedColumnFormula);
        Assert.AreEqual(formula, sheet.Cells["G2"].Formula);
        Assert.AreEqual(formula, sheet.Cells["G3"].Formula);
        Assert.AreEqual(formula, sheet.Cells["G11"].Formula);

        table.AddRow(3);
        Assert.AreEqual(formula, sheet.Cells["G13"].Formula);


        SaveAndCleanup(package);
    }
    [TestMethod]
    public void RenameTableWithCalculatedColumnFormulas()
    {
        using ExcelPackage? p = new ExcelPackage();
        // Get the worksheet containing the tables
        ExcelWorksheet? ws1 = p.Workbook.Worksheets.Add("Sheet1");
        ExcelWorksheet? ws2 = p.Workbook.Worksheets.Add("Sheet2");

        // Get the tables and check the calculated column formulas
        ExcelTable? tbl1 = ws1.Tables.Add(ws1.Cells["A1:C2"], "Table1");
        tbl1.Columns[2].CalculatedColumnFormula = "Table1[Column1]+Table1[Column2]";

        ExcelTable? tbl2 = ws1.Tables.Add(ws1.Cells["E1:G2"], "Table2");
        tbl2.Columns[2].CalculatedColumnFormula = "Table1[[#This Row],[Column1]]+Table2[[#This Row],[Column2]]";

        ws2.SetFormula(1, 1, "Table1[[#This Row],[Column1]]");
        ws2.Cells["B1:B2"].Formula = "Table1[[#This Row],[Column3]]";
        p.Workbook.Names.AddFormula("TableRef", "Table1[[#This Row],[Column1]]");
        Assert.AreEqual("Table1[Column1]+Table1[Column2]", tbl1.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("Table1[[#This Row],[Column1]]+Table2[[#This Row],[Column2]]", tbl2.Columns["Column3"].CalculatedColumnFormula);

        // Rename Table1 to Table3 and check the formulas were updated
        tbl1.Name = "NewTableName";
        Assert.AreEqual("NewTableName[Column1]+NewTableName[Column2]", tbl1.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("NewTableName[[#This Row],[Column1]]+Table2[[#This Row],[Column2]]", tbl2.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("NewTableName[[#This Row],[Column1]]", p.Workbook.Worksheets[1].Cells["A1"].Formula);
        Assert.AreEqual("NewTableName[[#This Row],[Column3]]", p.Workbook.Worksheets[1].Cells["B2"].Formula);
        Assert.AreEqual("NewTableName[[#This Row],[Column1]]", p.Workbook.Names["TableRef"].Formula);
    }
    [TestMethod]
    public void RenameTableWithCalculatedColumnFormulasSameStartOfTableName()
    {
        using ExcelPackage? p = new ExcelPackage();
        // Create some worksheets
        ExcelWorksheet? ws1 = p.Workbook.Worksheets.Add("Sheet1");
        ExcelWorksheet? ws2 = p.Workbook.Worksheets.Add("Sheet2");

        // Create some tables with calculated column formulas
        ExcelTable? tbl1 = ws1.Tables.Add(ws1.Cells["A1:C2"], "Table1");
        tbl1.Columns[2].CalculatedColumnFormula = "Table1[Column1]+Table1[Column2]";

        ExcelTable? tbl2 = ws1.Tables.Add(ws1.Cells["E1:G2"], "Table12");
        tbl2.Columns[2].CalculatedColumnFormula = "Table1[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]";

        // Create some references outside of the table
        ws2.SetFormula(1, 1, "Table1[[#This Row],[Column1]]");
        ws2.Cells["B1:B2"].Formula = "Table1[[#This Row],[Column3]]";
        p.Workbook.Names.AddFormula("TableRef", "Table1[[#This Row],[Column1]]");
        Assert.AreEqual("Table1[Column1]+Table1[Column2]", tbl1.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("Table1[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]", tbl2.Columns["Column3"].CalculatedColumnFormula);
        Assert.AreEqual("Table1[Column1]+Table1[Column2]", ws1.Cells["C2"].Formula);
        Assert.AreEqual("Table1[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]", ws1.Cells["G2"].Formula);

        // Rename Table1 to Table3 and check the formulas were updated
        tbl1.Name = "Table3";
        Assert.AreEqual("Table3[Column1]+Table3[Column2]", tbl1.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("Table3[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]", tbl2.Columns[2].CalculatedColumnFormula);
        Assert.AreEqual("Table3[Column1]+Table3[Column2]", ws1.Cells["C2"].Formula);
        Assert.AreEqual("Table3[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]", ws1.Cells["G2"].Formula);
        Assert.AreEqual("Table3[[#This Row],[Column1]]", p.Workbook.Worksheets[1].Cells["A1"].Formula);
        Assert.AreEqual("Table3[[#This Row],[Column3]]", p.Workbook.Worksheets[1].Cells["B2"].Formula);
        Assert.AreEqual("Table3[[#This Row],[Column1]]", p.Workbook.Names["TableRef"].Formula);
    }
    [TestMethod]
    public void CalculatedColumnFormula_SetToEmptyString()
    {
        using ExcelPackage? pck = new ExcelPackage();
        // Set up a worksheet containing a table
        ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
        wks.Cells["A1"].Value = "Col1";
        wks.Cells["B1"].Value = "Col2";
        wks.Cells["C1"].Value = "Col3";
        wks.Cells["A2"].Value = 1;
        wks.Cells["B2"].Value = 2;
        ExcelTable? table1 = wks.Tables.Add(wks.Cells["A1:C2"], "Table1");
        string? formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
        table1.Columns[2].CalculatedColumnFormula = formula;

        // Check the calculated column formula
        Assert.AreEqual(formula, wks.Cells["C2"].Formula);
        Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

        // Remove the calculated column formula from the table
        table1.Columns["Col3"].CalculatedColumnFormula = null;

        // Check the formula has been removed from the table
        Assert.IsTrue(string.IsNullOrEmpty(wks.Cells["C2"].Formula));
        Assert.IsTrue(string.IsNullOrEmpty(table1.Columns["Col3"].CalculatedColumnFormula));

        pck.SaveAs(@"C:\epplusTest\Testoutput\CalculatedColumnFormula_SetToEmptyString.xlsx");

        // NOW OPEN THE FILE IN EXCEL - IS IT CORRUPT?
        Assert.Inconclusive();
    }

    [TestMethod]
    public void CalculatedColumnFormula_RemoveFormulas()
    {
        using ExcelPackage? p = OpenPackage("CalculatedColumnFormulaRemove1.xlsx", true);
        // Set up a worksheet containing a table
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Col1";
        ws.Cells["B1"].Value = "Col2";
        ws.Cells["C1"].Value = "Col3";
        ws.Cells["A2"].Value = 1;
        ws.Cells["B2"].Value = 2;
        ExcelTable? table1 = ws.Tables.Add(ws.Cells["A1:C2"], "Table1");
        string? formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
        table1.Columns[2].CalculatedColumnFormula = formula;

        // Check the calculated column formula
        Assert.AreEqual(formula, ws.Cells["C2"].Formula);
        Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

        // Remove all formulas from the table
        table1.Range.ClearFormulas();
        table1.Range.ClearFormulaValues();

        // Check the calculated column formula is no longer there
        Assert.IsTrue(string.IsNullOrEmpty(table1.Columns["Col3"].CalculatedColumnFormula));
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void CalculatedColumnFormula_RemoveFormulas_AddRow()
    {
        using ExcelPackage? p = OpenPackage("CalculatedColumnFormulaRemove2.xlsx", true);
        // Set up a worksheet containing a table
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Col1";
        ws.Cells["B1"].Value = "Col2";
        ws.Cells["C1"].Value = "Col3";
        ws.Cells["A2"].Value = 1;
        ws.Cells["B2"].Value = 2;
        ExcelTable? table1 = ws.Tables.Add(ws.Cells["A1:C2"], "Table1");
        string? formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
        table1.Columns[2].CalculatedColumnFormula = formula;

        // Check the calculated column formula
        Assert.AreEqual(formula, ws.Cells["C2"].Formula);
        Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

        // Remove all formulas from the table
        table1.Range.ClearFormulas();
        table1.Range.ClearFormulaValues();
        Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["C2"].Formula));

        // Add a row to the table
        table1.InsertRow(1);

        // Check the formula has not been reinserted
        Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["C2"].Formula));
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void CalculatedColumnFormula_OneCellDifferent_AddRow()
    {
        using ExcelPackage? p = OpenPackage("CalculatedColumnFormulaRemove3.xlsx", true);
        // Set up a worksheet containing a table
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Col1";
        ws.Cells["B1"].Value = "Col2";
        ws.Cells["C1"].Value = "Col3";
        ws.Cells["A2"].Value = 1;
        ws.Cells["B2"].Value = 2;
        ws.Cells["A3"].Value = 3;
        ws.Cells["B3"].Value = 4;
        ws.Cells["A4"].Value = 5;
        ws.Cells["B4"].Value = 6;
        ExcelTable? table1 = ws.Tables.Add(ws.Cells["A1:C4"], "Table1");
        string? formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
        table1.Columns[2].CalculatedColumnFormula = formula;

        // Check the calculated column formula has been added to each cell
        Assert.AreEqual(formula, ws.Cells["C2"].Formula);
        Assert.AreEqual(formula, ws.Cells["C3"].Formula);
        Assert.AreEqual(formula, ws.Cells["C4"].Formula);
        Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

        // Remove the calculated column formula from one row and use a different formula instead
        ws.Cells["C3"].ClearFormulas();
        ws.Cells["C3"].ClearFormulaValues();
        string? differentFormula = "Table1[[#This Row],[Col1]]";
        ws.Cells["C3"].Formula = differentFormula;
        Assert.AreEqual(differentFormula, ws.Cells["C3"].Formula);

        // Add a new row to the bottom of the table
        table1.AddRow();

        // Check that the new row has the formula
        Assert.AreEqual(formula, ws.Cells["C5"].Formula);
        Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

        // Check the cell where we used a different formula hasn't changed
        Assert.AreEqual(differentFormula, ws.Cells["C3"].Formula);
        SaveAndCleanup(p);
    }
    [TestMethod]
    public void CreateTableAfterDeletingAMergedCell()
    {
        // Reproduce issue 780
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? worksheet = package.Workbook.Worksheets.Add("Sheet1");

        // Prepare some data
        worksheet.Cells["A1"].Value = "Column 1";
        worksheet.Cells["A2"].Value = 1;
        worksheet.Cells["B1"].Value = "Column 2";
        worksheet.Cells["B2"].Value = 2;

        // Merge cells in row 4 (not related to the data above)
        worksheet.Cells["A4:B4"].Merge = true;
        // Delete the row that has the merged cells
        worksheet.DeleteRow(4);

        // Create a table
        ExcelRange? tableCells = worksheet.Cells["A1:B2"];
        ExcelTable? table = worksheet.Tables.Add(tableCells, "table"); // --> This triggers a NullReferenceException
    }
}