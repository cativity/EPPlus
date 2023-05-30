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
using OfficeOpenXml.Table;
using System;

namespace EPPlusTest.Table;

[TestClass]
public class TableInsertTests : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("TableInsert.xlsx", true);
    }

    [ClassCleanup]
    public static void Cleanup() => SaveAndCleanup(_pck);

    #region Insert Row

    [TestMethod]
    public void TableInsertRowTop()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableInsertTop");
        LoadTestdata(ws, 100);

        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertTop");
        ws.Cells["A102"].Value = "Shift Me Down";
        _ = tbl.InsertRow(0);

        Assert.AreEqual("A1:D101", tbl.Address.Address);
        Assert.IsNull(tbl.Range.Offset(1, 0, 1, 1).Value);
        Assert.IsNull(tbl.Range.Offset(1, 1, 1, 1).Value);
        Assert.IsNull(tbl.Range.Offset(1, 2, 1, 1).Value);
        Assert.IsNull(tbl.Range.Offset(1, 3, 1, 1).Value);
        Assert.IsNull(tbl.Range.Offset(1, 4, 1, 1).Value);
        Assert.AreEqual("Shift Me Down", ws.Cells["A103"].Value);
        _ = tbl.InsertRow(0, 3);
        Assert.AreEqual("Shift Me Down", ws.Cells["A106"].Value);
    }

    [TestMethod]
    public void TableInsertRowBottom()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableInsertBottom");
        LoadTestdata(ws, 100);
        ws.Cells["A102"].Value = "Shift Me Down";
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertBottom");
        _ = tbl.AddRow(1);
        Assert.AreEqual("A1:D101", tbl.Address.Address);
        Assert.AreEqual("Shift Me Down", ws.Cells["A103"].Value);
        _ = tbl.AddRow(3);
        Assert.AreEqual("A1:D104", tbl.Address.Address);
        Assert.AreEqual("Shift Me Down", ws.Cells["A106"].Value);
    }

    [TestMethod]
    public void TableInsertRowBottomWithTotal()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableInsertBottomTotal");
        LoadTestdata(ws, 100, 2);
        ws.Cells["B102"].Value = "Shift Me Down";
        ws.Cells["F5"].Value = "Don't Shift Me";

        ExcelTable? tbl = ws.Tables.Add(ws.Cells["B1:E100"], "TableInsertBottomTotal");
        tbl.ShowTotal = true;
        tbl.Columns[0].TotalsRowFunction = RowFunctions.Sum;
        tbl.Columns[1].TotalsRowFunction = RowFunctions.Count;
        tbl.Columns[2].TotalsRowFunction = RowFunctions.Average;
        tbl.Columns[3].TotalsRowFunction = RowFunctions.CountNums;
        _ = tbl.AddRow(1);
        Assert.AreEqual("B1:E102", tbl.Address.Address);
        Assert.AreEqual("Shift Me Down", ws.Cells["B103"].Value);
        _ = tbl.AddRow(3);
        Assert.AreEqual("B1:E105", tbl.Address.Address);
        Assert.AreEqual("Don't Shift Me", ws.Cells["F5"].Value);
        Assert.AreEqual("Shift Me Down", ws.Cells["B106"].Value);
    }

    [TestMethod]
    public void TableInsertRowInside()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableInsertRowInside");
        LoadTestdata(ws, 100);
        ws.Cells["A102"].Value = "Shift Me Down";
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertRowInside");
        _ = tbl.InsertRow(98);
        Assert.AreEqual("A1:D101", tbl.Address.Address);
        Assert.AreEqual("Shift Me Down", ws.Cells["A103"].Value);
        _ = tbl.InsertRow(1, 3);
        Assert.AreEqual("A1:D104", tbl.Address.Address);
        Assert.AreEqual("Shift Me Down", ws.Cells["A106"].Value);
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void TableInsertRowPositionNegative()
    {
        //Setup
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Table1");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table1");
        _ = tbl.InsertRow(-1);
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void TableInsertRowRowsNegative()
    {
        //Setup
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Table1");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table1");
        _ = tbl.InsertRow(0, -1);
    }

    [TestMethod]
    public void TableAddRowToMax()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableMaxRow");
        LoadTestdata(ws, 100);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableMaxRow");

        //Act
        _ = tbl.AddRow(ExcelPackage.MaxRows - 100);

        //Assert
        Assert.AreEqual(ExcelPackage.MaxRows, tbl.Address._toRow);
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidOperationException))]
    public void TableAddRowOverMax()
    {
        using ExcelPackage? p = new ExcelPackage();

        //Setup
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("TableOverMaxRow");
        LoadTestdata(ws, 100);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableOverMaxRow");

        //Act
        _ = tbl.AddRow(ExcelPackage.MaxRows - 99);

        //Assert
        Assert.AreEqual(ExcelPackage.MaxRows, tbl.Address._toRow);
    }

    #endregion

    #region Insert Column

    [TestMethod]
    public void TableInsertColumnFirst()
    {
        //Setup

        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableInsertColFirst");
        LoadTestdata(ws, 100);

        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertColFirst");
        ws.Cells["E10"].Value = "Shift Me Right";
        _ = tbl.Columns.Insert(0);

        Assert.AreEqual("A1:E100", tbl.Address.Address);
        Assert.IsNull(tbl.Range.Offset(2, 0, 1, 1).Value);
        Assert.IsNull(tbl.Range.Offset(100, 0, 1, 1).Value);
        Assert.AreEqual("Shift Me Right", ws.Cells["F10"].Value);
        Assert.AreEqual("Column1", tbl.Columns[0].Name);
        _ = tbl.Columns.Insert(0, 3);
        Assert.AreEqual("Shift Me Right", ws.Cells["I10"].Value);
    }

    [TestMethod]
    public void TableAddColumn()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableAddCol");
        LoadTestdata(ws, 100);
        ws.Cells["E99"].Value = "Shift Me Right";
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableAddColumn");
        _ = tbl.Columns.Add(1);
        Assert.AreEqual("A1:E100", tbl.Address.Address);
        Assert.AreEqual("Shift Me Right", ws.Cells["F99"].Value);
        _ = tbl.Columns.Add(3);
        Assert.AreEqual("A1:H100", tbl.Address.Address);
        Assert.AreEqual("Shift Me Right", ws.Cells["I99"].Value);
    }

    [TestMethod]
    public void TableAddColumnWithTotal()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableAddColTotal");
        LoadTestdata(ws, 100, 2);
        ws.Cells["F100"].Value = "Shift Me Right";
        ws.Cells["A50,F102"].Value = "Don't Shift Me";

        ExcelTable? tbl = ws.Tables.Add(ws.Cells["B1:E100"], "TableAddTotal");
        tbl.ShowTotal = true;
        tbl.Columns[0].TotalsRowFunction = RowFunctions.Sum;
        tbl.Columns[1].TotalsRowFunction = RowFunctions.Count;
        tbl.Columns[2].TotalsRowFunction = RowFunctions.Average;
        tbl.Columns[3].TotalsRowFunction = RowFunctions.CountNums;
        _ = tbl.Columns.Insert(0, 1);
        Assert.AreEqual("B1:F101", tbl.Address.Address);
        Assert.AreEqual(RowFunctions.Sum, tbl.Columns[1].TotalsRowFunction);
        Assert.AreEqual(RowFunctions.CountNums, tbl.Columns[4].TotalsRowFunction);
        Assert.AreEqual("Shift Me Right", ws.Cells["G100"].Value);
        _ = tbl.Columns.Add(3);
        Assert.AreEqual("B1:I101", tbl.Address.Address);
        Assert.AreEqual("Don't Shift Me", ws.Cells["A50"].Value);
        Assert.AreEqual("Don't Shift Me", ws.Cells["F102"].Value);
    }

    [TestMethod]
    public void TableInsertColumnInside()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableInsertColInside");
        LoadTestdata(ws, 100);
        ws.Cells["E9999"].Value = "Don't Me Down";
        ws.Cells["E19999"].Value = "Don't Me Down";
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertColInside");
        _ = tbl.Columns.Insert(4, 2);
        Assert.AreEqual("A1:F100", tbl.Address.Address);
        _ = tbl.Columns.Insert(8, 8);
        Assert.AreEqual("A1:N100", tbl.Address.Address);
        Assert.AreEqual("Don't Me Down", ws.Cells["E9999"].Value);
        Assert.AreEqual("Don't Me Down", ws.Cells["E19999"].Value);
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void TableInsertColumnPositionNegative()
    {
        //Setup
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Table1");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table1");
        _ = tbl.Columns.Insert(-1);
    }

    [TestMethod]
    public void TableAddColumnToMax()
    {
        using ExcelPackage? p = new ExcelPackage(); // We discard this as it takes to long time to save

        //Setup
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("TableMaxColumn");
        LoadTestdata(ws, 100);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableMaxColumn");

        //Act
        _ = tbl.Columns.Add(ExcelPackage.MaxColumns - 4);

        //Assert
        Assert.AreEqual(ExcelPackage.MaxColumns, tbl.Address._toCol);
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidOperationException))]
    public void TableAddColumnOverMax()
    {
        using ExcelPackage? p = new ExcelPackage();

        //Setup
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("TableOverMaxColumn");
        LoadTestdata(ws, 100);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableOverMaxRow");

        //Act
        _ = tbl.Columns.Add(ExcelPackage.MaxColumns - 3);
    }

    #endregion

    [TestMethod]
    public void AddRowsToTablesOfDifferentWidths_TopWider()
    {
        using ExcelPackage? pck = OpenTemplatePackage("TestTableAddRows.xlsx");

        // Get sheet 1 from the workbook, and get the tables we are going to test
        ExcelWorksheet? ws = TryGetWorksheet(pck, "Sheet1");
        ExcelTable? table1 = ws.Tables["Table1"];
        ExcelTable? table2 = ws.Tables["Table2"];

        // Make sure the tables are where we expect them to be
        if (table1.Address.ToString() != "B2:E3")
        {
            Assert.Inconclusive();
        }

        if (table2.Address.ToString() != "B6:C7")
        {
            Assert.Inconclusive();
        }

        // Add 10 rows to Table1
        _ = table1.AddRow(10);

        // Make sure Table1's address has been correctly updated
        Assert.AreEqual("B2:E13", table1.Address.ToString());

        // Make sure Table2 below has been correctly moved
        Assert.AreEqual("B16:C17", table2.Address.ToString());

        // Add 10 rows to Table2
        _ = table2.AddRow(10);

        // Make sure Table2 has been correctly updated
        Assert.AreEqual("B16:C27", table2.Address.ToString());

        // Make sure Table1 hasn't moved
        Assert.AreEqual("B2:E13", table1.Address.ToString());
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidOperationException))]
    public void AddRowsToTablesOfDifferentWidths_BottomWider()
    {
        using ExcelPackage? pck = OpenTemplatePackage("TestTableAddRows.xlsx");

        // Get sheet 2 from the workbook, and get the tables we are going to test
        ExcelWorksheet? ws = TryGetWorksheet(pck, "Sheet2");
        ExcelTable? table3 = ws.Tables["Table3"];
        ExcelTable? table4 = ws.Tables["Table4"];

        // Make sure the tables are where we expect them to be
        if (table3.Address.ToString() != "B2:C3")
        {
            Assert.Inconclusive();
        }

        if (table4.Address.ToString() != "B6:E7")
        {
            Assert.Inconclusive();
        }

        // Add 10 rows to Table3
        _ = table3.AddRow(10);

        // Make sure Table3's address has been correctly updated
        Assert.AreEqual("B2:C13", table3.Address.ToString());

        // Make sure Table4 below has been correctly moved
        Assert.AreEqual("B16:E17", table4.Address.ToString());

        // Add 10 rows to Table4
        _ = table4.AddRow(10);

        // Make sure Table4 has been correctly updated
        Assert.AreEqual("B16:E27", table4.Address.ToString());

        // Make sure Table3 hasn't moved
        Assert.AreEqual("B2:C13", table3.Address.ToString());
    }

    [TestMethod]
    public void TableAddOneColumnStartingFromA()
    {
        using ExcelPackage? p = OpenPackage("TestTableAdd1Column.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:A10"], "Table1");
        _ = tbl.Columns.Add(1);
        Assert.AreEqual("A1:B10", tbl.Address.Address);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void TableInsertAddRowShowHeaderFalse()
    {
        using ExcelPackage? p = OpenPackage("TableAddRowWithoutHeader.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:A10"], "Table1");
        tbl.ShowHeader = false;

        _ = tbl.InsertRow(0, 1);
        Assert.AreEqual("A1:A11", tbl.Address.Address);

        _ = tbl.InsertRow(1, 1);
        Assert.AreEqual("A1:A12", tbl.Address.Address);

        _ = tbl.AddRow(1);
        Assert.AreEqual("A1:A13", tbl.Address.Address);
    }

    [TestMethod]
    public void TableInsertAddColumnShowHeaderFalse()
    {
        using ExcelPackage? p = OpenPackage("TableAddColWithoutHeader.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:A10"], "Table1");
        tbl.ShowHeader = false;

        _ = tbl.InsertColumn(0, 1);
        Assert.AreEqual("A1:B10", tbl.Address.Address);

        _ = tbl.InsertColumn(1, 1);
        Assert.AreEqual("A1:C10", tbl.Address.Address);
    }

    [TestMethod]
    public void TableDeleteRowShowHeaderFalse()
    {
        using ExcelPackage? p = OpenPackage("TableDeleteRowWithoutHeader.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:A10"], "Table1");
        tbl.ShowHeader = false;

        _ = tbl.DeleteRow(0, 3);
        Assert.AreEqual("A1:A7", tbl.Address.Address);

        _ = tbl.DeleteRow(1, 2);
        Assert.AreEqual("A1:A5", tbl.Address.Address);
    }

    [TestMethod]
    public void TableWithCalculatedFormulaInsert()
    {
        using ExcelPackage? p = OpenPackage("TableCalculatedColumnInsert.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[2, 1, 10, 1].Value = 1;
        ws.Cells[2, 2, 10, 2].Value = 2;
        ws.Cells[2, 3, 10, 3].Value = 3;
        ws.Cells[2, 4, 10, 4].Value = 4;
        ws.Cells[2, 5, 10, 5].Value = 5;
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:E10"], "Table1");
        tbl.ShowHeader = true;
        tbl.Columns[2].CalculatedColumnFormula = "Table1[[#This Row],[Column1]]";
        Assert.AreEqual("Table1[[#This Row],[Column1]]", ws.Cells["C2"].Formula);
        Assert.AreEqual("", ws.Cells["D11"].Formula);
        _ = tbl.Columns.Insert(1, 1);
        _ = tbl.InsertRow(0, 2);
        _ = tbl.Columns.Add(3);
        _ = tbl.AddRow(5);
        Assert.AreEqual("Table1[[#This Row],[Column1]]", ws.Cells["D2"].Formula);
        Assert.AreEqual("Table1[[#This Row],[Column1]]", ws.Cells["D12"].Formula);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void TableWithCalculatedFormulaDelete()
    {
        using ExcelPackage? p = OpenPackage("TableCalculatedColumnDelete.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[2, 1, 10, 1].Value = 1;
        ws.Cells[2, 2, 10, 2].Value = 2;
        ws.Cells[2, 3, 10, 3].Value = 3;
        ws.Cells[2, 4, 10, 4].Value = 4;
        ws.Cells[2, 5, 10, 5].Value = 5;
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:E10"], "Table1");
        tbl.ShowHeader = true;
        tbl.Columns[2].CalculatedColumnFormula = "Table1[[#This Row],[Column1]]";
        Assert.AreEqual("Table1[[#This Row],[Column1]]", ws.Cells["C2"].Formula);
        Assert.AreEqual("", ws.Cells["D11"].Formula);
        _ = tbl.Columns.Delete(1, 1);
        _ = tbl.DeleteRow(1, 2);
        Assert.AreEqual("Table1[[#This Row],[Column1]]", ws.Cells["B2"].Formula);
        Assert.AreEqual("", ws.Cells["B9"].Formula);
        SaveAndCleanup(p);
    }
}