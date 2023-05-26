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
using System.Globalization;
using System.Threading;

namespace EPPlusTest.Filter
{
    [TestClass]
    public class ValueFilter : TestBase
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
        public void ValuesFilter()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueFilter");
            LoadTestdata(ws);
            
            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            ExcelValueFilterColumn? col=ws.AutoFilter.Columns.AddValueFilterColumn(1);
            col.Filters.Add("7");
            col.Filters.Add("14");
            col.Filters.Add("88");
            col.Filters.Add("sss");
            col.Filters.Blank = true;
            col.Filters.Add(new ExcelFilterDateGroupItem(2018, 12));
            col.Filters.Add(new ExcelFilterDateGroupItem(2019, 1, 15));

            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(6).Hidden);
            Assert.AreEqual(false, ws.Row(7).Hidden);
            Assert.AreEqual(true, ws.Row(8).Hidden);
            Assert.AreEqual(true, ws.Row(13).Hidden);
            Assert.AreEqual(false, ws.Row(88).Hidden);
            Assert.AreEqual(true, ws.Row(100).Hidden);
            Assert.AreEqual(false, ws.Row(101).Hidden);
        }
        [TestMethod]
        public void DateFilterYear()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateYear");
            LoadTestdata(ws, 200);

            ws.AutoFilterAddress = ws.Cells["A1:D200"];
            ExcelValueFilterColumn? col = ws.AutoFilter.Columns.AddValueFilterColumn(0);
            int year = DateTime.Today.Year - 1;
            col.Filters.Add(new ExcelFilterDateGroupItem(year));
            ws.AutoFilter.ApplyFilter();

            int row = GetRowFromDate(new DateTime(year, 12, 15));
            Assert.AreEqual(false, ws.Row(row).Hidden);
            row = GetRowFromDate(new DateTime(year, 12, 31));
            Assert.AreEqual(false, ws.Row(row).Hidden);
            row = GetRowFromDate(new DateTime(year+1, 1, 1));
            Assert.AreEqual(true, ws.Row(row).Hidden);
        }
        [TestMethod]
        public void DateFilterMonth()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateMonth");
            LoadTestdata(ws, 200);

            ws.AutoFilterAddress = ws.Cells["A1:D200"];
            ExcelValueFilterColumn? col = ws.AutoFilter.Columns.AddValueFilterColumn(0);
            int year = DateTime.Today.Year;
            col.Filters.Add(new ExcelFilterDateGroupItem(year,1));
            ws.AutoFilter.ApplyFilter();

            int row = GetRowFromDate(new DateTime(year-1, 12, 31));
            Assert.AreEqual(true, ws.Row(row).Hidden);
            row = GetRowFromDate(new DateTime(year, 1, 1));
            Assert.AreEqual(false, ws.Row(row).Hidden);
            row = GetRowFromDate(new DateTime(year, 1, 31));
            Assert.AreEqual(false, ws.Row(row).Hidden);
            row = GetRowFromDate(new DateTime(year, 2, 1));
            Assert.AreEqual(true, ws.Row(row).Hidden);
        }
        [TestMethod]
        public void DateFilterDay()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateDay");
            LoadTestdata(ws, 200);

            ws.AutoFilterAddress = ws.Cells["A1:D200"];
            ExcelValueFilterColumn? col = ws.AutoFilter.Columns.AddValueFilterColumn(0);
            int year = DateTime.Today.Year;
            col.Filters.Add(new ExcelFilterDateGroupItem(year, 1, 12));
            ws.AutoFilter.ApplyFilter();

            int row = GetRowFromDate(new DateTime(year, 1, 11));
            Assert.AreEqual(true, ws.Row(row).Hidden);
            row = GetRowFromDate(new DateTime(year, 1, 12));
            Assert.AreEqual(false, ws.Row(row).Hidden);
            row = GetRowFromDate(new DateTime(year, 1, 13));
            Assert.AreEqual(true, ws.Row(row).Hidden);
        }
        [TestMethod]
        public void DateFilterHour()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateHour");
            LoadTestdata(ws, 200);
            int year = DateTime.Today.Year;
            ws.SetValue("A82", new DateTime(year, 1, 20, 12, 11, 33));
            ws.SetValue("A83", new DateTime(year, 1, 20, 13, 11, 33));
            ws.AutoFilterAddress = ws.Cells["A1:D200"];
            ExcelValueFilterColumn? col = ws.AutoFilter.Columns.AddValueFilterColumn(0);
            col.Filters.Add(new ExcelFilterDateGroupItem(year, 1, 20, 12));
            ws.AutoFilter.ApplyFilter();

            int row = GetRowFromDate(new DateTime(year, 1, 19));
            Assert.AreEqual(true, ws.Row(row).Hidden);
            row++;
            Assert.AreEqual(false, ws.Row(row).Hidden);
            row++;
            Assert.AreEqual(true, ws.Row(row).Hidden);
        }
        [TestMethod]
        public void DateFilterMinute()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateMinute");
            LoadTestdata(ws, 200);
            int year = DateTime.Today.Year;
            ws.SetValue("A82", new DateTime(year, 1, 20, 12, 11, 33));
            ws.SetValue("A83", new DateTime(year, 1, 20, 12, 12, 33));
            ws.AutoFilterAddress = ws.Cells["A1:D200"];
            ExcelValueFilterColumn? col = ws.AutoFilter.Columns.AddValueFilterColumn(0);
            col.Filters.Add(new ExcelFilterDateGroupItem(year, 1, 20, 12, 11));
            ws.AutoFilter.ApplyFilter();

            int row = GetRowFromDate(new DateTime(year, 1, 19));
            Assert.AreEqual(true, ws.Row(row).Hidden);
            row++;
            Assert.AreEqual(false, ws.Row(row).Hidden);
            row++;
            Assert.AreEqual(true, ws.Row(row).Hidden);
        }
        [TestMethod]
        public void DateFilterSecond()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateSecond");
            LoadTestdata(ws, 200);
            int year = DateTime.Today.Year;
            ws.SetValue("A82", new DateTime(year, 1, 20, 12, 11, 33));
            ws.SetValue("A83", new DateTime(year, 1, 20, 12, 11, 35));
            ws.AutoFilterAddress = ws.Cells["A1:D200"];
            ExcelValueFilterColumn? col = ws.AutoFilter.Columns.AddValueFilterColumn(0);
            col.Filters.Add(new ExcelFilterDateGroupItem(year, 1, 20, 12, 11, 33));
            ws.AutoFilter.ApplyFilter();

            int row = GetRowFromDate(new DateTime(year, 1, 19));
            Assert.AreEqual(true, ws.Row(row).Hidden);
            row++;
            Assert.AreEqual(false, ws.Row(row).Hidden);
            row++;
            Assert.AreEqual(true, ws.Row(row).Hidden);
        }
        [TestMethod]
        public void TextFilter()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Text");
            LoadTestdata(ws);
            SetDateValues(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D102"];
            ExcelValueFilterColumn? col = ws.AutoFilter.Columns.AddValueFilterColumn(2);
            col.Filters.Add("Value 8");
            col.Filters.Add("Value 55");
            col.Filters.Add("Value 33");
            col.Filters.Blank = true;
            col.Filters.Add(new ExcelFilterDateGroupItem(2018, 12));
            col.Filters.Add(new ExcelFilterDateGroupItem(2019, 1, 15));

            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(7).Hidden);
            Assert.AreEqual(false, ws.Row(8).Hidden);
            Assert.AreEqual(true, ws.Row(9).Hidden);
            Assert.AreEqual(true, ws.Row(54).Hidden);
            Assert.AreEqual(false, ws.Row(55).Hidden);
            Assert.AreEqual(true, ws.Row(100).Hidden);
            Assert.AreEqual(false, ws.Row(101).Hidden); //Verify blanks
            Assert.AreEqual(false, ws.Row(102).Hidden); //Verify blanks
            Assert.AreEqual(false, ws.Row(103).Hidden);
        }
        [TestMethod]
        public void NumericFormattedFilter()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NumericFormatted");
            CultureInfo? currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-SE");
            LoadTestdata(ws);
            SetDateValues(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D102"];
            ExcelValueFilterColumn? col = ws.AutoFilter.Columns.AddValueFilterColumn(3);
            col.Filters.Add("66,00");
            col.Filters.Add("3 003,00");
            col.Filters.Add("3 036,00");
            col.Filters.Add("3 069,00");
            col.Filters.Add("3 102,00");
            col.Filters.Add("3 135,00");
            col.Filters.Add("3 168,00");
            col.Filters.Blank = true;

            ws.AutoFilter.ApplyFilter();
            Assert.AreEqual(false, ws.Row(2).Hidden);
            Assert.AreEqual(true, ws.Row(3).Hidden);
            Assert.AreEqual(true, ws.Row(90).Hidden);
            Assert.AreEqual(false, ws.Row(91).Hidden);
            Assert.AreEqual(false, ws.Row(92).Hidden);
            Assert.AreEqual(false, ws.Row(93).Hidden);
            Assert.AreEqual(false, ws.Row(94).Hidden);
            Assert.AreEqual(false, ws.Row(95).Hidden);
            Assert.AreEqual(false, ws.Row(96).Hidden);
            Assert.AreEqual(true, ws.Row(97).Hidden);

            Thread.CurrentThread.CurrentCulture= currentCulture;
        }

    }
}
