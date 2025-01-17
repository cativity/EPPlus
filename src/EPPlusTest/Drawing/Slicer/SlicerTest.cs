﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System.IO;
using OfficeOpenXml.Table;

namespace EPPlusTest.Drawing.Slicer;

[TestClass]
public class SlicerTest : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context) => _pck = OpenPackage("Slicer.xlsx", true);

    [ClassCleanup]
    public static void Cleanup()
    {
        string? dirName = _pck.File.DirectoryName;
        string? fileName = _pck.File.FullName;

        SaveAndCleanup(_pck);

        if (File.Exists(fileName))
        {
            File.Copy(fileName, dirName + "\\SlicerRead.xlsx", true);
        }
    }

    [TestMethod]
    public void AddTableSlicerDate()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableSlicerDate");

        LoadTestdata(ws);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table1");
        ExcelTableSlicer? slicer = ws.Drawings.AddTableSlicer(tbl.Columns[0]);

        _ = slicer.FilterValues.Add(new ExcelFilterDateGroupItem(2019, 11, 4));
        _ = slicer.FilterValues.Add(new ExcelFilterDateGroupItem(2019, 11, 5));
        _ = slicer.FilterValues.Add(new ExcelFilterDateGroupItem(2019, 11, 7));
        _ = slicer.FilterValues.Add(new ExcelFilterDateGroupItem(2019, 12));
        slicer.Cache.HideItemsWithNoData = true;
        slicer.SetPosition(1, 0, 5, 0);
        slicer.SetSize(200, 600);

        Assert.AreEqual(eCrossFilter.None, slicer.Cache.CrossFilter);
        Assert.IsTrue(slicer.Cache.HideItemsWithNoData);
        slicer.Cache.HideItemsWithNoData = false; //Validate element is removed
        Assert.IsFalse(slicer.Cache.HideItemsWithNoData);
        slicer.Cache.HideItemsWithNoData = true; //Add element again

        ExcelTableSlicer? slicer2 = ws.Drawings.AddTableSlicer(tbl.Columns[2]);
        slicer2.Style = eSlicerStyle.Light4;
        slicer2.LockedPosition = true;
        slicer2.ColumnCount = 3;
        slicer2.Cache.CrossFilter = eCrossFilter.None;
        slicer2.Cache.SortOrder = eSortOrder.Descending;
        slicer2.Cache.CustomListSort = false;

        slicer2.SetPosition(1, 0, 9, 0);
        slicer2.SetSize(200, 600);

        Assert.AreEqual(eCrossFilter.None, slicer2.Cache.CrossFilter);
        Assert.AreEqual(eSortOrder.Descending, slicer2.Cache.SortOrder);
        Assert.IsFalse(slicer2.Cache.CustomListSort);

        Assert.IsTrue(slicer2.LockedPosition);
        Assert.AreEqual(3, slicer2.ColumnCount);
        Assert.AreEqual("SlicerStyleLight4", slicer2.StyleName);
        Assert.AreEqual(eSlicerStyle.Light4, slicer2.Style);
    }

    [TestMethod]
    public void AddTableSlicerString()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableSlicerNumber");

        LoadTestdata(ws);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table2");
        ExcelTableSlicer? slicer = ws.Drawings.AddTableSlicer(tbl.Columns[1]);

        slicer.Style = eSlicerStyle.Dark1;
        _ = slicer.FilterValues.Add("52");
        _ = slicer.FilterValues.Add("53");
        _ = slicer.FilterValues.Add("61");
        _ = slicer.FilterValues.Add("102");
        slicer.StartItem = 50;
        slicer.ShowCaption = false;
        slicer.SetPosition(1, 0, 5, 0);
        slicer.SetSize(200, 600);

        Assert.AreEqual(50, slicer.StartItem);
        Assert.AreEqual(eSlicerStyle.Dark1, slicer.Style);
        Assert.IsFalse(slicer.ShowCaption);
    }

    [TestMethod]
    public void AddPivotTableSlicer()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotTableSlicer");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        rf.Items.Refresh();
        rf.Items[0].Hidden = true;
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;

        ExcelPivotTableSlicer? slicer = ws.Drawings.AddPivotTableSlicer(tbl.Fields[1]);
        slicer.Cache.Data.Items[0].Hidden = true;
        slicer.Cache.Data.Items[2].Hidden = true;
        slicer.Cache.Data.Items[4].Hidden = true;
        slicer.Style = eSlicerStyle.Light5;
        slicer.SetPosition(1, 0, 10, 0);
        slicer.SetSize(200, 600);
    }

    [TestMethod]
    public void AddPivotTableSlicerToTwoPivotTables()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("SlicerPivotSameCache");
        LoadTestdata(ws);
        ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Pivot1");
        _ = p1.RowFields.Add(p1.Fields[0]);

        _ = p1.DataFields.Add(p1.Fields[3]);
        ExcelPivotTable? p2 = ws.PivotTables.Add(ws.Cells["K1"], p1.CacheDefinition, "Pivot2");
        _ = p2.DataFields.Add(p2.Fields[1]);
        _ = p2.RowFields.Add(p2.Fields[3]);

        //p2.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days);

        ExcelPivotTableSlicer? slicer = ws.Drawings.AddPivotTableSlicer(p1.Fields[0]);
        slicer.Cache.PivotTables.Add(p2);
        p2.CacheDefinition.Refresh();
        slicer.Cache.Data.Items[0].Hidden = true;
        slicer.Cache.Data.Items[1].Hidden = true;
        slicer.Cache.Data.SortOrder = eSortOrder.Descending;
        slicer.Style = eSlicerStyle.Light5;
        slicer.SetPosition(1, 0, 15, 0);
        slicer.SetSize(200, 600);

        Assert.AreEqual(slicer.Cache.Data.SortOrder, eSortOrder.Descending);
        Assert.AreEqual(slicer.Style, eSlicerStyle.Light5);
        Assert.IsTrue(slicer.Cache.Data.Items[0].Hidden);
        Assert.IsTrue(slicer.Cache.Data.Items[1].Hidden);

        Assert.AreEqual(100, p1.Fields[0].Items.Count);
        Assert.IsTrue(p1.Fields[0].Items[0].Hidden);
        Assert.IsTrue(p1.Fields[0].Items[1].Hidden);

        Assert.AreEqual(100, p2.Fields[0].Items.Count);
        Assert.IsTrue(p2.Fields[0].Items[0].Hidden);
        Assert.IsTrue(p2.Fields[0].Items[1].Hidden);
    }

    [TestMethod]
    public void AddPivotTableSlicerToTwoPivotTablesWithDateGrouping()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("SlicerPivotSameCacheDateGroup");
        LoadTestdata(ws);
        ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Pivot1");
        _ = p1.RowFields.Add(p1.Fields[0]);

        _ = p1.DataFields.Add(p1.Fields[3]);
        ExcelPivotTable? p2 = ws.PivotTables.Add(ws.Cells["K1"], p1.CacheDefinition, "Pivot2");
        _ = p2.DataFields.Add(p2.Fields[1]);
        _ = p2.RowFields.Add(p2.Fields[3]);

        p1.Fields[0].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days);
        p1.Fields[0].Name = "Days";
        ExcelPivotTableSlicer? slicer = ws.Drawings.AddPivotTableSlicer(p1.Fields[0]);
        slicer.Cache.PivotTables.Add(p2);
        p2.CacheDefinition.Refresh();
        slicer.Cache.Data.Items[0].Hidden = true;
        slicer.Cache.Data.Items[1].Hidden = true;
        slicer.Cache.Data.SortOrder = eSortOrder.Ascending;
        slicer.Style = eSlicerStyle.Light4;
        slicer.SetPosition(1, 0, 15, 0);
        slicer.SetSize(200, 600);

        Assert.AreEqual(slicer.Cache.Data.SortOrder, eSortOrder.Ascending);
        Assert.AreEqual(slicer.Style, eSlicerStyle.Light4);
        Assert.IsTrue(slicer.Cache.Data.Items[0].Hidden);
        Assert.IsTrue(slicer.Cache.Data.Items[1].Hidden);

        Assert.AreEqual(369, p1.Fields[0].Items.Count);
        Assert.IsTrue(p1.Fields[0].Items[0].Hidden);
        Assert.IsTrue(p1.Fields[0].Items[1].Hidden);

        Assert.AreEqual(369, p2.Fields[0].Items.Count);
        Assert.IsTrue(p2.Fields[0].Items[0].Hidden);
        Assert.IsTrue(p2.Fields[0].Items[1].Hidden);
    }

    [TestMethod]
    public void AddPivotTableSlicerToTwoPivotTablesWithNumberGrouping()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("SlicerPivotSameCacheNumberGroup");
        LoadTestdata(ws);
        ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Pivot1");
        _ = p1.RowFields.Add(p1.Fields[1]);

        _ = p1.DataFields.Add(p1.Fields[3]);
        ExcelPivotTable? p2 = ws.PivotTables.Add(ws.Cells["K1"], p1.CacheDefinition, "Pivot2");
        _ = p2.DataFields.Add(p2.Fields[1]);
        _ = p2.RowFields.Add(p2.Fields[3]);

        p1.Fields[1].AddNumericGrouping(0, 100, 5);
        ExcelPivotTableSlicer? slicer = ws.Drawings.AddPivotTableSlicer(p1.Fields[1]);
        slicer.Cache.PivotTables.Add(p2);
        p2.CacheDefinition.Refresh();
        slicer.Cache.Data.Items[0].Hidden = true;
        slicer.Cache.Data.Items[1].Hidden = true;
        slicer.Cache.Data.SortOrder = eSortOrder.Descending;
        slicer.Style = eSlicerStyle.Light5;
        slicer.SetPosition(1, 0, 15, 0);
        slicer.SetSize(200, 600);
    }

    [TestMethod]
    public void RemovePivotTableSlicerIfLast()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("RemovedPivotTableSlicerLast");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        rf.Items.Refresh();
        rf.Items[0].Hidden = true;
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;

        ExcelPivotTableSlicer? slicer = ws.Drawings.AddPivotTableSlicer(tbl.Fields[1]);
        Assert.AreEqual(slicer, tbl.Fields[1].Slicer);

        ws.Drawings.Remove(slicer);
        Assert.IsNull(tbl.Fields[1].Slicer);
    }

    [TestMethod]
    public void RemoveOnePivotTableSlicerNotLast()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("RemovedPivotTableSlicerNotLast");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        rf.Items.Refresh();
        rf.Items[0].Hidden = true;
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;

        ExcelPivotTableSlicer? slicer1 = ws.Drawings.AddPivotTableSlicer(tbl.Fields[1]);
        ExcelPivotTableSlicer? slicer2 = ws.Drawings.AddPivotTableSlicer(tbl.Fields[2]);
        Assert.AreEqual(slicer1, tbl.Fields[1].Slicer);
        Assert.AreEqual(slicer2, tbl.Fields[2].Slicer);

        ws.Drawings.Remove(slicer1);
        Assert.IsNull(tbl.Fields[1].Slicer);
        Assert.IsNotNull(tbl.Fields[2].Slicer);
    }

    [TestMethod]
    public void RemoveTableSlicerStringIfLast()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableSlicerRemoveLast");

        LoadTestdata(ws);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table3");
        ExcelTableSlicer? slicer = ws.Drawings.AddTableSlicer(tbl.Columns[1]);

        ws.Drawings.Remove(slicer);
        Assert.IsNull(tbl.Columns[1].Slicer);
    }

    [TestMethod]
    public void RemoveTableSlicerStringIfNotLast()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableSlicerRemoveNotLast");

        LoadTestdata(ws);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table4");
        ExcelTableSlicer? slicer1 = ws.Drawings.AddTableSlicer(tbl.Columns[1]);
        _ = ws.Drawings.AddTableSlicer(tbl.Columns[2]);

        ws.Drawings.Remove(slicer1);
        Assert.IsNull(tbl.Columns[1].Slicer);
        Assert.IsNotNull(tbl.Columns[2].Slicer);
    }

    [TestMethod]
    public void AddTwoTableSlicersToSameColumn()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableSlicerTwoOnSameColumn");

        LoadTestdata(ws);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table5");
        ExcelTableSlicer? slicer1 = ws.Drawings.AddTableSlicer(tbl.Columns[2]);
        ExcelTableSlicer? slicer2 = ws.Drawings.AddTableSlicer(tbl.Columns[2]);
        slicer1.SetPosition(0, 0, 5, 0);
        slicer2.SetPosition(11, 0, 5, 0);
        _ = slicer1.FilterValues.Add("Value 12");
        _ = slicer1.FilterValues.Add("value 10");
        _ = slicer2.FilterValues.Add("value 15");
        slicer2.StartItem = 2;
        Assert.AreEqual(2, ws.Drawings.Count);
        Assert.IsNull(tbl.Columns[1].Slicer);
        Assert.AreEqual(slicer1, tbl.Columns[2].Slicer);
    }

    [TestMethod]
    public void AddTwoPivotTableSlicersToSameField()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PTSlicerTwoOnSameField");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        rf.Items.Refresh();
        rf.Items[0].Hidden = true;
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;

        ExcelPivotTableSlicer? slicer1 = ws.Drawings.AddPivotTableSlicer(tbl.Fields[2]);
        ExcelPivotTableSlicer? slicer2 = ws.Drawings.AddPivotTableSlicer(tbl.Fields[2]);
        slicer1.SetPosition(0, 0, 5, 0);
        slicer2.SetPosition(11, 0, 5, 0);
        Assert.AreEqual(slicer1, tbl.Fields[2].Slicer);
    }
}