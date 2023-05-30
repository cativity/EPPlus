using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable;

[TestClass]
public class PivotTableCaptionFilterTests : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("PivotTableFilters.xlsx", true);
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Data1");
        ExcelRangeBase? r = LoadItemData(ws);
        _ = ws.Tables.Add(r, "Table1");
        ws = _pck.Workbook.Worksheets.Add("Data2");
        r = LoadItemData(ws);
        _ = ws.Tables.Add(r, "Table2");
    }

    [ClassCleanup]
    public static void Cleanup() => SaveAndCleanup(_pck);

    [TestMethod]
    public void AddCaptionEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionEqual, "Hardware");
    }

    [TestMethod]
    public void AddCaptionNotEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionNotEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotEqual, "Hardware");
    }

    [TestMethod]
    public void AddCaptionBeginsWithFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionBeginsWith");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionBeginsWith, "H");
    }

    [TestMethod]
    public void AddCaptionNotBeginsWithFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionNotBeginsWith");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotBeginsWith, "H");
    }

    [TestMethod]
    public void AddCaptionEndsWithFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionEndsWith");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionEndsWith, "ware");
    }

    [TestMethod]
    public void AddCaptionNotEndsWithFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionNotEndsWith");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotEndsWith, "ware");
    }

    [TestMethod]
    public void AddCaptionContainsFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionContains");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionContains, "roc");
    }

    [TestMethod]
    public void AddCaptionNotContainsFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionNotContains");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotContains, "roc");
    }

    [TestMethod]
    public void AddCaptionGreaterThanFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionGreaterThan");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionGreaterThan, "H");
    }

    [TestMethod]
    public void AddCaptionGreaterThanOrEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionGreaterThanOrEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionGreaterThanOrEqual, "H");
    }

    [TestMethod]
    public void AddCaptionLessThanFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionLessThan");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionLessThan, "H");
    }

    [TestMethod]
    public void AddCaptionLessThanOrEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionLessThanOrEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionLessThanOrEqual, "H");
    }

    [TestMethod]
    public void AddCaptionBetweenWithFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionBetween");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionBetween, "H", "I");

        Assert.AreEqual("H", pt.Fields[1].Filters[0].StringValue1);
        Assert.AreEqual("I", pt.Fields[1].Filters[0].StringValue2);
        Assert.AreEqual(2, ((ExcelCustomFilterColumn)pt.Fields[1].Filters[0].Filter).Filters.Count);
    }

    [TestMethod]
    public void AddCaptionNotBetweenWithFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaptionNotBetween");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);

        _ = pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotBetween, "H", "I");
        Assert.AreEqual("H", pt.Fields[1].Filters[0].StringValue1);
        Assert.AreEqual("I", pt.Fields[1].Filters[0].StringValue2);
        Assert.AreEqual(2, ((ExcelCustomFilterColumn)pt.Fields[1].Filters[0].Filter).Filters.Count);
    }
}