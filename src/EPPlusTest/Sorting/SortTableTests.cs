﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Sorting;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Sorting;

[TestClass]
public class SortTableTests
{
    private static ExcelTable CreateTable(ExcelWorksheet sheet, bool addTotalsRow = true)
    {
        // header
        sheet.Cells[1, 1].Value = "Header1";
        sheet.Cells[1, 2].Value = "Header2";
        sheet.Cells[1, 3].Value = "Header3";

        // row 1
        sheet.Cells[2, 1].Value = 10;
        sheet.Cells[2, 2].Value = 2;
        sheet.Cells[2, 3].Value = 3;

        // row 2
        sheet.Cells[3, 1].Value = 5;
        sheet.Cells[3, 2].Value = 2;
        sheet.Cells[3, 3].Value = 3;

        ExcelTable? table = sheet.Tables.Add(sheet.Cells["A1:C3"], "myTable");
        table.TableStyle = TableStyles.Dark1;
        table.ShowTotal = addTotalsRow;
        table.Columns[0].TotalsRowFunction = RowFunctions.Sum;
        table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
        table.Columns[2].TotalsRowFunction = RowFunctions.Sum;

        return table;
    }

    [TestMethod]
    public void TableSortByColumnIndex()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        ExcelTable? table = CreateTable(sheet);

        table.Sort(x => x.SortBy.Column(0));

        Assert.AreEqual(5, sheet.Cells[2, 1].Value);
        Assert.AreEqual(10, sheet.Cells[3, 1].Value);
        Assert.IsNotNull(table.SortState, "SortState was null");
        Assert.IsNotNull(table.SortState.SortConditions, "SortState.SortConditions was null");
        Assert.IsFalse(table.SortState.SortConditions.First().Descending, "First SortCondition was not descending");
    }

    [TestMethod]
    public void TableSortByColumnName()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        ExcelTable? table = CreateTable(sheet);

        table.Sort(x => x.SortBy.ColumnNamed("Header1"));
        Assert.AreEqual(5, sheet.Cells[2, 1].Value);
        Assert.AreEqual(10, sheet.Cells[3, 1].Value);
        Assert.IsNotNull(table.SortState, "SortState was null");
        Assert.IsNotNull(table.SortState.SortConditions, "SortState.SortConditions was null");
        Assert.IsFalse(table.SortState.SortConditions.First().Descending, "First SortCondition was not descending");
    }

    [TestMethod]
    public void TableSortByCustomList()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

        // header
        sheet.Cells[1, 1].Value = "Size";
        sheet.Cells[1, 2].Value = "Price";
        sheet.Cells[1, 3].Value = "Color";

        // row 1
        sheet.Cells[2, 1].Value = "M";
        sheet.Cells[2, 2].Value = 20;
        sheet.Cells[2, 3].Value = "Blue";

        // row 2
        sheet.Cells[3, 1].Value = "XL";
        sheet.Cells[3, 2].Value = 25;
        sheet.Cells[3, 3].Value = "Yellow";

        // row 3
        sheet.Cells[4, 1].Value = "S";
        sheet.Cells[4, 2].Value = 10;
        sheet.Cells[4, 3].Value = "Yellow";

        // row 4
        sheet.Cells[5, 1].Value = "L";
        sheet.Cells[5, 2].Value = 21;
        sheet.Cells[5, 3].Value = "Blue";

        // row 5
        sheet.Cells[6, 1].Value = "S";
        sheet.Cells[6, 2].Value = 20;
        sheet.Cells[6, 3].Value = "Blue";

        // row 6
        sheet.Cells[7, 1].Value = "S";
        sheet.Cells[7, 2].Value = 10;
        sheet.Cells[7, 3].Value = "Blue";

        ExcelTable? table = sheet.Tables.Add(sheet.Cells["A1:C7"], "myTable");

        table.Sort(x => x.SortBy.ColumnNamed("Size")
                         .UsingCustomList("S", "M", "L", "XL")
                         .ThenSortBy.ColumnNamed("Price", eSortOrder.Descending)
                         .ThenSortBy.Column(2)
                         .UsingCustomList("Blue", "Yellow"));

        Assert.AreEqual("S", sheet.Cells[2, 1].Value, $"First row, first col not 'S' but '{sheet.Cells[2, 1].Value}'");
        Assert.AreEqual(20, sheet.Cells[2, 2].Value, $"First row, second col not 20 but '{sheet.Cells[2, 2].Value}'");
        Assert.AreEqual("Blue", sheet.Cells[2, 3].Value, $"First row, third col not 'Blue' but '{sheet.Cells[2, 1].Value}'");

        Assert.AreEqual("S", sheet.Cells[3, 1].Value);
        Assert.AreEqual(10, sheet.Cells[3, 2].Value);
        Assert.AreEqual("Blue", sheet.Cells[3, 3].Value);

        Assert.AreEqual("S", sheet.Cells[4, 1].Value);
        Assert.AreEqual(10, sheet.Cells[4, 2].Value);
        Assert.AreEqual("Yellow", sheet.Cells[4, 3].Value);

        Assert.AreEqual("M", sheet.Cells[5, 1].Value);
        Assert.AreEqual("L", sheet.Cells[6, 1].Value);
        Assert.AreEqual("XL", sheet.Cells[7, 1].Value);

        //package.SaveAs(new FileInfo(@"c:\Temp\TableSort2.xlsx"));
    }

    [TestMethod]
    public void SortShouldRetainRelativeTableAddresses()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("Sheet1");
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
        ws.Cells["C2"].Value = 11;
        ws.Cells["C3"].Value = 10;
        ws.Cells["C4"].Value = 9;
        ws.Cells["C5"].Value = 8;
        ws.Cells["C6"].Value = 7;
        ws.Cells["C7"].Value = 6;
        ws.Cells["C8"].Value = 5;
        ws.Cells["C9"].Value = 4;
        ws.Cells["C10"].Value = 3;
        ws.Cells["C11"].Value = 2;
        ws.Cells["C12"].Value = 1;

        tbl.Sort(x => x.SortBy.Column(1, eSortOrder.Ascending));
        Assert.AreEqual("TestTable[[#This Row],[123]]+TestTable[[#This Row],[Column2]]", tbl.Columns[14].CalculatedColumnFormula);
    }
}