using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Table.PivotTable;

[TestClass]
public class PivotTableValueFilterTests : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("PivotTableValueFilters.xlsx", true);
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
    public void AddValueEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueEqual, 0, 5);
    }

    [TestMethod]
    public void AddValueNotEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueNotEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);
        pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueNotEqual, 0, 12.2);

        //pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueNotEqual, 0, 85.2);
    }

    [TestMethod]
    public void AddValueGreaterThanFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueGreaterThan");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[3].Filters.AddValueFilter(ePivotTableValueFilterType.ValueGreaterThan, 0, 12.2);
        pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueLessThan, 0, 500);
    }

    [TestMethod]
    public void AddValueGreaterThanOrEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueGreaterThanOrEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueGreaterThanOrEqual, 0, 12.2);
    }

    [TestMethod]
    public void AddValueLessThanFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueLessThan");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueLessThan, 0, 12.2);
    }

    [TestMethod]
    public void AddValueLessThanOrEqualFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueLessThanOrEqual");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueLessThanOrEqual, 0, 12.2);
    }

    [TestMethod]
    public void AddValueBetweenFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueBetweeen");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueBetween, 0, 4, 10);
    }

    [TestMethod]
    public void AddValueNotBetweenFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueNotBetweeen");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueNotBetween, df, 4, 10);
    }

    [TestMethod]
    public void AddTop10CountFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueTop15Count");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Count, df, 15);
    }

    [TestMethod]
    public void AddTop10PercentFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueTop20Percent");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Percent, df, 20);
    }

    [TestMethod]
    public void AddTop10SumFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueTop25Sum");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Sum, df, 25);
    }

    [TestMethod]
    public void AddBottom10CountFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueBottom15Count");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Count, df, 15, false);
    }

    [TestMethod]
    public void AddBottom10PercentFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueBottom20Percent");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Percent, df, 20, false);
    }

    [TestMethod]
    public void AddBottom10SumFilter()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data1"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ValueBottom25Sum");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
        pt.RowFields.Add(pt.Fields[4]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);

        pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Sum, df, 25, false);
        ws.Cells["B4:D4"].Merge = true;
    }
}