
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System;

namespace EPPlusTest.Table.PivotTable;

[TestClass]
public class PivotTableDateFilterTests : TestBase
{
    static ExcelPackage _pck;
    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("PivotTableDateFilters.xlsx", true);
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Data1");
        ExcelRangeBase? r = LoadItemData(ws);
        ws.Tables.Add(r, "Table1");
        ws = _pck.Workbook.Worksheets.Add("Data2");
        r = LoadItemData(ws);
        ws.Tables.Add(r, "Table2");
    }
    [ClassCleanup]
    public static void Cleanup()
    {
        SaveAndCleanup(_pck);
    }
    [TestMethod]
    public void AddDateEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateEqual, new DateTime(2010,3,31));
    }
    [TestMethod]
    public void AddDateNotEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateNotEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNotEqual, new DateTime(2010, 3, 31));
    }
    [TestMethod]
    public void AddDateOlderFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateBefore");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateOlderThan, new DateTime(2010, 3, 31));
    }
    [TestMethod]
    public void AddDateOlderOrEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateBeforeOrEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateOlderThanOrEqual, new DateTime(2010, 3, 31));
    }
    [TestMethod]
    public void AddDateNewerFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateNewer");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNewerThan, new DateTime(2010, 3, 31));
    }
    [TestMethod]
    public void AddDateNewerOrEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateNewerOrEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNewerThanOrEqual, new DateTime(2010, 3, 31));
    }
    [TestMethod]
    public void AddDateBetweenFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateBetween");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateBetween, new DateTime(2010, 3, 31), new DateTime(2010, 6, 30));
    }
    [TestMethod]
    public void AddDateNotBetweenFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DateNotBetween");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNotBetween, new DateTime(2010, 3, 31), new DateTime(2010, 6, 30));
    }
    [TestMethod]
    public void AddDateLastMonthFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LastMonth");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.LastMonth);
    }
    [TestMethod]
    public void AddDateLastQuarterFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LastQuarter");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.LastQuarter);
    }
    [TestMethod]
    public void AddDateLastWeekFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LastWeek");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.LastWeek);
    }
    [TestMethod]
    public void AddDateLastYearFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LastYear");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.LastYear);
    }
    [TestMethod]
    public void AddDateM1Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M1");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M1);
    }
    [TestMethod]
    public void AddDateM2Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M2");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M2);
    }
    [TestMethod]
    public void AddDateM3Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M3");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M3);
    }
    [TestMethod]
    public void AddDate42Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M4");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M4);
    }
    [TestMethod]
    public void AddDateM5Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M5");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M5);
    }
    [TestMethod]
    public void AddDateM6Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M6");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M6);
    }
    [TestMethod]
    public void AddDateM7Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M7");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M7);
    }
    [TestMethod]
    public void AddDateM8Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M8");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M8);
    }
    [TestMethod]
    public void AddDateM9Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M9");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M9);
    }
    [TestMethod]
    public void AddDateM10Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M10");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M10);
    }
    [TestMethod]
    public void AddDateM11Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M11");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M11);
    }
    [TestMethod]
    public void AddDateM12Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("M12");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M12);
    }
    [TestMethod]
    public void AddDateQ1Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Q1");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Q1);
    }
    [TestMethod]
    public void AddDateQ2Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Q2");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Q2);
    }
    [TestMethod]
    public void AddDateQ3Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Q3");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Q3);
    }
    [TestMethod]
    public void AddDateQ4Filter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Q4");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Q4);
    }
    [TestMethod]
    public void AddDateYesterdayFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Yesterday");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Yesterday);
    }
    [TestMethod]
    public void AddDateTodayFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Today");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Today);
    }
    [TestMethod]
    public void AddDateTomorrowFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Tomorrow");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Tomorrow);
    }
    [TestMethod]
    public void AddDateYTDFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("YTD");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.YearToDate);
    }
    [TestMethod]
    public void AddDateThisMonthFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ThisMonth");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.ThisMonth);
    }
    [TestMethod]
    public void AddDateThisQuarterFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ThisQuarter");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.ThisQuarter);
    }
    [TestMethod]
    public void AddDateThisWeekFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ThisWeek");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.ThisWeek);
    }
    [TestMethod]
    public void AddDateThisYearFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ThisYear");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.ThisYear);
    }
    [TestMethod]
    public void AddDateNextMonthFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NextMonth");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.NextMonth);
    }
    [TestMethod]
    public void AddDateNextQuarterFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NextQuarter");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.NextQuarter);
    }
    [TestMethod]
    public void AddDateNextWeekFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NextWeek");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.NextWeek);
    }
    [TestMethod]
    public void AddDateNextYearFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("NextYear");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.NextYear);
    }
}