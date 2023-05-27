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
using System;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Filter;

namespace EPPlusTest.Filter;

[TestClass]
public class DynamicFilterTest : TestBase
{
    static ExcelPackage _pck;
    [ClassInitialize]
    public static void Init(TestContext context)
    {
        _pck = OpenPackage("ValueFilter.xlsx", true);
    }
    [ClassCleanup]
    public static void Cleanup()
    {
        SaveAndCleanup(_pck);
    }
    [TestMethod]
    public void AboveAverage()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("AboveAverage");
        LoadTestdata(ws);
        SetDateValues(ws);

        ws.AutoFilterAddress = ws.Cells["A1:D100"];
        ExcelDynamicFilterColumn? col=ws.AutoFilter.Columns.AddDynamicFilterColumn(1);
        col.Type = eDynamicFilterType.AboveAverage;
        ws.AutoFilter.ApplyFilter();    
        Assert.AreEqual(true, ws.Row(48).Hidden);
        Assert.AreEqual(false, ws.Row(50).Hidden);
        Assert.AreEqual(false, ws.Row(51).Hidden);
        Assert.AreEqual(false, ws.Row(52).Hidden);
        Assert.AreEqual(true, ws.Row(53).Hidden);
    }
    [TestMethod]
    public void BelowAverage()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("BelowAverage");
        LoadTestdata(ws);
        SetDateValues(ws);

        ws.AutoFilterAddress = ws.Cells["A1:D100"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(1);
        col.Type = eDynamicFilterType.BelowAverage;
        ws.AutoFilter.ApplyFilter();
        Assert.AreEqual(false, ws.Row(48).Hidden);
        Assert.AreEqual(true, ws.Row(50).Hidden);
        Assert.AreEqual(true, ws.Row(51).Hidden);
        Assert.AreEqual(true, ws.Row(52).Hidden);
        Assert.AreEqual(false, ws.Row(53).Hidden);
    }
    #region Day
    [TestMethod]
    public void Yesterday()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Yesterday");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.Yesterday;
        ws.AutoFilter.ApplyFilter();
            
        //Assert
        DateTime dt = DateTime.Today.AddDays(-1);
        int row = GetRowFromDate(dt);
        Assert.AreEqual(true, ws.Row(row - 1).Hidden);
        Assert.AreEqual(false, ws.Row(row).Hidden);
        Assert.AreEqual(true, ws.Row(row + 1).Hidden);
    }
    [TestMethod]
    public void Today()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Today");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.Today;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime dt = DateTime.Today;
        int row = GetRowFromDate(dt);
        Assert.AreEqual(true, ws.Row(row - 1).Hidden);
        Assert.AreEqual(false, ws.Row(row).Hidden);
        Assert.AreEqual(true, ws.Row(row + 1).Hidden);
    }
    [TestMethod]
    public void Tomorrow()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Tomorrow");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.Tomorrow;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime dt = DateTime.Today.AddDays(1);
        int row = GetRowFromDate(dt);
        Assert.AreEqual(true, ws.Row(row - 1).Hidden);
        Assert.AreEqual(false, ws.Row(row).Hidden);
        Assert.AreEqual(true, ws.Row(row + 1).Hidden);
    }

    #endregion
    #region Week
    [TestMethod]
    public void LastWeek()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LastWeek");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.LastWeek;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = GetPrevSunday(DateTime.Today.AddDays(-7));
        int startRow = GetRowFromDate(dt);
        int endRow = startRow + 6;
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void ThisWeek()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ThisWeek");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.ThisWeek;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime dt = GetPrevSunday(DateTime.Today);
        int startRow = GetRowFromDate(dt);
        int endRow = startRow + 6;
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void NextWeek()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NextWeek");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.NextWeek;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime dt = GetPrevSunday(DateTime.Today.AddDays(7));
        int startRow = GetRowFromDate(dt);
        int endRow = startRow + 6;
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    #endregion
    #region Month
    [TestMethod]
    public void LastMonth()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LastMonth");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.LastMonth;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(-1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1).AddMonths(1).AddDays(-1));
        Assert.AreEqual(true, ws.Row(startRow-1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow+1).Hidden);
    }
    [TestMethod]
    public void ThisMonth()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ThisMonth");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.ThisMonth;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today;
        int startRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1).AddMonths(1).AddDays(-1));
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void NextMonth()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NextMonth");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.NextMonth;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1).AddMonths(1).AddDays(-1));
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void M1()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M1");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M1;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 1, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 1, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

    }
    [TestMethod]
    public void M2()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M2");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M2;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 2, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 2, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

    }
    [TestMethod]
    public void M3()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M3");
        LoadTestdata(ws, 600);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D600"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M3;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 3, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 3, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

    }
    [TestMethod]
    public void M4()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M4");
        LoadTestdata(ws, 600);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D600"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M4;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 4, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 4, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

    }
    [TestMethod]
    public void M5()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M5");
        LoadTestdata(ws, 700);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D700"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M5;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 5, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 5, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

    }
    [TestMethod]
    public void M6()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M6");
        LoadTestdata(ws, 700);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D700"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M6;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 6, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 6, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void M7()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M7");
        LoadTestdata(ws, 700);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D700"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M7;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 7, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 7, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void M8()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M8");
        LoadTestdata(ws, 800);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D800"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M8;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 8, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 8, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void M9()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M9");
        LoadTestdata(ws, 800);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D800"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M9;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 9, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 9, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void M10()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M10");
        LoadTestdata(ws, 800);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D800"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M10;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 10, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 10, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void M11()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M11");
        LoadTestdata(ws, 800);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D800"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M11;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 11, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 11, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void M12()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M12");
        LoadTestdata(ws, 900);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D900"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.M12;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today.AddMonths(1);
        int startRow = GetRowFromDate(new DateTime(dt.Year, 12, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 12, 1).AddMonths(1).AddDays(-1));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }

    #endregion
    #region Quarter
    [TestMethod]
    public void LastQuarter()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LastQuarter");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.LastQuarter;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime startDate = GetStartOfQuarter(DateTime.Today.AddMonths(-3));
        DateTime endDate = startDate.AddMonths(3).AddDays(-1);
        int startRow = GetRowFromDate(startDate);
        int endRow = GetRowFromDate(endDate);
        if (startRow > 2)
        {
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        }
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void ThisQuarter()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ThisQuarter");
        LoadTestdata(ws, 500);

        //Act   
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.ThisQuarter;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime startDate = GetStartOfQuarter(DateTime.Today);
        DateTime endDate = startDate.AddMonths(3).AddDays(-1);
        int startRow = GetRowFromDate(startDate);
        int endRow = GetRowFromDate(endDate);
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void NextQuarter()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NextQuarter");
        LoadTestdata(ws, 600);

        //Act   
        ws.AutoFilterAddress = ws.Cells["A1:D600"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.NextQuarter;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime startDate = GetStartOfQuarter(DateTime.Today.AddMonths(3));
        DateTime endDate = startDate.AddMonths(3).AddDays(-1);
        int startRow = GetRowFromDate(startDate);
        int endRow = GetRowFromDate(endDate);
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void Q1()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Q1");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.Q1;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today;
        int startRow = GetRowFromDate(new DateTime(dt.Year, 1, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 3, 31));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void Q2()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Q2");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.Q2;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today;
        int startRow = GetRowFromDate(new DateTime(dt.Year, 4, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 6, 30));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void Q3()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Q3");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.Q3;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today;
        int startRow = GetRowFromDate(new DateTime(dt.Year, 7, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 9, 30));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void Q4()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Q4");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.Q4;
        ws.AutoFilter.ApplyFilter();


        //Assert
        DateTime dt = DateTime.Today;
        int startRow = GetRowFromDate(new DateTime(dt.Year, 10, 1));
        int endRow = GetRowFromDate(new DateTime(dt.Year, 12, 31));
        //Will only verify this year
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }

    #endregion
    #region Year
    [TestMethod]
    public void LastYear()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LastYear");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.LastYear;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime startDate = new DateTime(DateTime.Today.Year-1, 1, 1);
        DateTime endDate = new DateTime(DateTime.Today.Year - 1, 12, 31);
        int startRow = GetRowFromDate(startDate);
        int endRow = GetRowFromDate(endDate);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    [TestMethod]
    public void ThisYear()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ThisYear");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.ThisYear;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime startDate = new DateTime(DateTime.Today.Year , 1, 1);
        DateTime endDate = new DateTime(DateTime.Today.Year, 12, 31);
        int startRow = GetRowFromDate(startDate);
        int endRow = GetRowFromDate(endDate);
        Assert.AreEqual(true, ws.Row(startRow-1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }

    [TestMethod]
    public void NextYear()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NextYear");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.NextYear;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime startDate = new DateTime(DateTime.Today.Year + 1, 1, 1);
        DateTime endDate = new DateTime(DateTime.Today.Year + 1, 12, 31);
        int startRow = GetRowFromDate(startDate);
        int endRow = GetRowFromDate(endDate);
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(500).Hidden);
        Assert.AreEqual(false, ws.Row(501).Hidden);
    }

    [TestMethod]
    public void YearToDate()
    {
        //Setup
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("YearToDate");
        LoadTestdata(ws, 500);

        //Act
        ws.AutoFilterAddress = ws.Cells["A1:D500"];
        ExcelDynamicFilterColumn? col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
        col.Type = eDynamicFilterType.YearToDate;
        ws.AutoFilter.ApplyFilter();

        //Assert
        DateTime startDate = new DateTime(DateTime.Today.Year, 1, 1);
        DateTime endDate = DateTime.Today;
        int startRow = GetRowFromDate(startDate);
        int endRow = GetRowFromDate(endDate);
        Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
        Assert.AreEqual(false, ws.Row(startRow).Hidden);
        Assert.AreEqual(false, ws.Row(endRow).Hidden);
        Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
    }
    #endregion

    #region Private methods
    private static DateTime GetStartOfQuarter(DateTime dt)
    {
        int quarter = (dt.Month - ((dt.Month - 1) % 3) + 1) / 3;
                      
        return new DateTime(dt.Year, (quarter * 3) + 1, 1);
    }
    private static DateTime GetPrevSunday(DateTime dt)
    {
        while (dt.DayOfWeek != DayOfWeek.Sunday)
        {
            dt = dt.AddDays(-1);
        }
        return dt;
    }
    #endregion
}