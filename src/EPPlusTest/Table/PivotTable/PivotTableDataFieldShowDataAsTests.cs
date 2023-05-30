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
  02/11/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlusTest.Table.PivotTable;

[TestClass]
public class PivotTableDataFieldShowDataAsTests : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("PivotTableShowDataAs.xlsx", true);
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Data");
        _ = LoadItemData(ws);
    }

    [ClassCleanup]
    public static void Cleanup() => SaveAndCleanup(_pck);

    [TestMethod]
    public void ShowAsPercentOfTotal()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercTot");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercTot");
        _ = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        df.ShowDataAs.SetPercentOfTotal();
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.PercentOfTotal, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsPercentOfRow()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercRow");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercRow");
        _ = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        df.ShowDataAs.SetPercentOfRow();
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.PercentOfRow, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsPercentOfCol()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercCol");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercCol");
        _ = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        df.ShowDataAs.SetPercentOfColumn();
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.PercentOfColumn, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsPercent()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPerc");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePerc");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        rf.Items.Refresh();
        df.ShowDataAs.SetPercent(rf, 50);
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.Percent, df.ShowDataAs.Value);
        Assert.AreEqual(rf.Index, df.BaseField);
        Assert.AreEqual(50, df.BaseItem);
    }

    [TestMethod]
    public void ShowAsIndex()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsIndex");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableIndex");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        rf.Items.Refresh();
        df.ShowDataAs.SetIndex();
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.Index, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsDifference()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsDifference");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableDifference");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        rf.Items.Refresh();
        df.ShowDataAs.SetDifference(rf, ePrevNextPivotItem.Previous);
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.Difference, df.ShowDataAs.Value);
        Assert.AreEqual(rf.Index, df.BaseField);
        Assert.AreEqual((int)ePrevNextPivotItem.Previous, df.BaseItem);
    }

    [TestMethod]
    public void ShowAsPercentageDifference()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercDiff");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercDiff");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        rf.Items.Refresh();
        df.ShowDataAs.SetPercentageDifference(rf, ePrevNextPivotItem.Previous);
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.PercentDifference, df.ShowDataAs.Value);
        Assert.AreEqual(rf.Index, df.BaseField);
        Assert.AreEqual((int)ePrevNextPivotItem.Previous, df.BaseItem);
    }

    [TestMethod]
    public void ShowAsRunningTotal()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsRunningTotal");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableRunningTotal");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        df.ShowDataAs.SetRunningTotal(rf);
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.RunningTotal, df.ShowDataAs.Value);
        Assert.AreEqual(rf.Index, df.BaseField);
    }

    [TestMethod]
    public void ShowAsPercentOfParent()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercentOfParent");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercParent");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        df.ShowDataAs.SetPercentParent(rf);
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.PercentOfParent, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsPercentOfParentRow()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercentOfParentRow");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercentOfParentRow");
        _ = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        df.ShowDataAs.SetPercentParentRow();
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.PercentOfParentRow, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsPercentOfParentCol()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercentOfParentCol");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercentOfParentRow");
        _ = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        df.ShowDataAs.SetPercentParentColumn();
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.PercentOfParentColumn, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsPercentOfRunningTotal()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercentOfRunningTotal");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercentOfParentRow");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.ShowDataAs.SetPercentOfRunningTotal(rf);
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.PercentOfRunningTotal, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsRankAscending()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsRankAscending");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableRankAscending");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.ShowDataAs.SetRankAscending(rf);
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.RankAscending, df.ShowDataAs.Value);
    }

    [TestMethod]
    public void ShowAsRankDescending()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ShowDataAsRankDescending");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableRankDescending");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.ShowDataAs.SetRankDescending(rf);
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;

        Assert.AreEqual(eShowDataAs.RankDescending, df.ShowDataAs.Value);
    }
}