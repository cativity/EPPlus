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
using System.Xml;

namespace EPPlusTest.Table.PivotTable;

[TestClass]
public class PivotTableTests : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("PivotTable.xlsx", true);
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Data");
        LoadItemData(ws);
    }

    [ClassCleanup]
    public static void Cleanup()
    {
        SaveAndCleanup(_pck);
    }

    [TestMethod]
    public void ValidateLoadSaveTableSource()
    {
        using ExcelPackage p1 = new ExcelPackage();
        string? tblName = "Table1";
        string? tblAddress = "A1:D4";
        ExcelWorksheet? wsData = p1.Workbook.Worksheets.Add("TableData");
        wsData.Cells["A1"].Value = "Column1";
        wsData.Cells["B1"].Value = "Column2";
        wsData.Cells["C1"].Value = "Column3";
        wsData.Cells["D1"].Value = "Column4";
        ExcelWorksheet? wsPivot = p1.Workbook.Worksheets.Add("PivotSimple");
        ExcelTable? Table1 = wsData.Tables.Add(wsData.Cells[tblAddress], tblName);
        ExcelPivotTable? pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], wsData.Cells[Table1.Address.Address], "PivotTable1");

        pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
        pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
        pivotTable1.ColumnFields.Add(pivotTable1.Fields[2]);

        Assert.AreEqual(tblAddress, wsPivot.PivotTables[0].CacheDefinition.SourceRange.Address);
        Assert.AreEqual(Table1.Columns.Count, pivotTable1.Fields.Count);
        Assert.AreEqual(1, pivotTable1.RowFields.Count);
        Assert.AreEqual(1, pivotTable1.DataFields.Count);
        Assert.AreEqual(1, pivotTable1.ColumnFields.Count);

        p1.Save();

        using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
        wsData = p2.Workbook.Worksheets[0];
        wsPivot = p2.Workbook.Worksheets[1];

        pivotTable1 = wsPivot.PivotTables[0];
        Assert.AreEqual(tblAddress, pivotTable1.CacheDefinition.SourceRange.Address);
        Assert.AreEqual(Table1.Columns.Count, pivotTable1.Fields.Count);
        Assert.AreEqual(1, pivotTable1.RowFields.Count);
        Assert.AreEqual(1, pivotTable1.DataFields.Count);
        Assert.AreEqual(1, pivotTable1.ColumnFields.Count);
    }

    [TestMethod]
    public void ValidateLoadSaveAddressSource()
    {
        using ExcelPackage p1 = new ExcelPackage();
        string? address = "A1:D4";
        ExcelWorksheet? wsData = p1.Workbook.Worksheets.Add("TableData");
        wsData.Cells["A1"].Value = "Column1";
        wsData.Cells["B1"].Value = "Column2";
        wsData.Cells["C1"].Value = "Column3";
        wsData.Cells["D1"].Value = "Column4";
        ExcelWorksheet? wsPivot = p1.Workbook.Worksheets.Add("PivotSimple");
        ExcelPivotTable? pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], wsData.Cells[address], "PivotTable1");
        pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
        pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
        pivotTable1.ColumnFields.Add(pivotTable1.Fields[2]);

        Assert.AreEqual(address, wsPivot.PivotTables[0].CacheDefinition.SourceRange.Address);
        Assert.AreEqual(4, pivotTable1.Fields.Count);
        Assert.AreEqual(1, pivotTable1.RowFields.Count);
        Assert.AreEqual(1, pivotTable1.DataFields.Count);
        Assert.AreEqual(1, pivotTable1.ColumnFields.Count);

        p1.Save();

        using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
        wsData = p2.Workbook.Worksheets[0];
        wsPivot = p2.Workbook.Worksheets[1];

        pivotTable1 = wsPivot.PivotTables[0];
        Assert.AreEqual(address, pivotTable1.CacheDefinition.SourceRange.Address);
        Assert.AreEqual(4, pivotTable1.Fields.Count);
        Assert.AreEqual(1, pivotTable1.RowFields.Count);
        Assert.AreEqual(1, pivotTable1.DataFields.Count);
        Assert.AreEqual(1, pivotTable1.ColumnFields.Count);
    }

    [TestMethod]
    public void CreatePivotTableAddressSource()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotSourceAddress");
        LoadTestdata(ws);

        ExcelPivotTable? pivotTable1 = ws.PivotTables.Add(ws.Cells["G1"], ws.Cells["A1:D100"], "PivotTable1");

        pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
        pivotTable1.RowFields.Add(pivotTable1.Fields[2]);
        pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
        pivotTable1.DataFields.Add(pivotTable1.Fields[3]);
    }

    [TestMethod]
    public void CreatePivotTableTableSource()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotSourceTable");
        LoadTestdata(ws);
        ExcelTable? table = ws.Tables.Add(ws.Cells["A1:D100"], "table1");
        ExcelPivotTable? pivotTable1 = ws.PivotTables.Add(ws.Cells["G1"], table, "PivotTable1");

        pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
        pivotTable1.RowFields.Add(pivotTable1.Fields[2]);
        pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
        pivotTable1.DataFields.Add(pivotTable1.Fields[3]);
    }

    [TestMethod]
    public void RowsDataOnColumns()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Rows-Data on columns");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable1");
        pt.GrandTotalCaption = "Total amount";
        pt.RowFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);
        pt.DataFields[0].Function = DataFieldFunctions.Product;
        pt.DataOnRows = false;
    }

    [TestMethod]
    public void RowsDataOnRow()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Rows-Data on rows");
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);
        pt.DataFields[0].Function = DataFieldFunctions.Average;
        pt.DataOnRows = true;
    }

    [TestMethod]
    public void ColumnsDataOnColumns()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Columns-Data on columns");
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
        pt.ColumnFields.Add(pt.Fields[1]);
        pt.ColumnFields.Add(pt.Fields[0]);
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);
        pt.DataOnRows = false;
    }

    [TestMethod]
    public void ColumnsDataOnRows()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Columns-Data on rows");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable4");
        pt.ColumnFields.Add(pt.Fields[1]);
        pt.ColumnFields.Add(pt.Fields[0]);
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);
        pt.DataOnRows = true;
    }

    [TestMethod]
    public void ColumnsRows_DataOnColumns()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on columns");
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable5");
        pt.ColumnFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);
        pt.DataOnRows = false;
    }

    [TestMethod]
    public void ColumnsRows_DataOnRows()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on rows");
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable6");
        pt.ColumnFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);
        pt.DataOnRows = true;
        ws.Drawings.AddChart("Pivotchart6", OfficeOpenXml.Drawing.Chart.eChartType.BarStacked3D, pt);
    }

    [TestMethod]
    public void RowsPage_DataOnColumns()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Rows/Page-Data on Columns");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:N11"], "Pivottable7");
        pt.PageFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);
        pt.DataOnRows = false;

        pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Max;
        Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Sum | eSubTotalFunctions.Max);

        pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP;
        Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP);

        pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.None;
        Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.None);

        pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Default;
        Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Default);

        pt.Fields[0].Sort = eSortType.Descending;
        pt.TableStyle = TableStyles.Medium14;
    }

    [TestMethod]
    public void Pivot_GroupDate()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pivot-Group Date");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:O11"], "Pivottable8");
        pt.RowFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[4]);

        pt.Fields[4]
          .AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days | eDateGroupBy.Quarters,
                           new DateTime(2010, 01, 31),
                           new DateTime(2010, 11, 30));

        pt.RowHeaderCaption = "År";
        pt.Fields[4].Name = "Dag";
        pt.Fields[4].Items[0].Hidden = true;
        pt.Fields[5].Name = "Månad";
        pt.Fields[5].Items[0].Hidden = true;
        pt.Fields[6].Name = "Kvartal";
        pt.Fields[6].Items[0].Hidden = true;
        pt.GrandTotalCaption = "Totalt";

        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);
        pt.DataOnRows = true;

        pt = ws.PivotTables.Add(ws.Cells["H3"], wsData.Cells["K1:O11"], "Pivottable10");
        pt.RowFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[4]);
        pt.Fields[4].AddDateGrouping(7, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
        pt.RowHeaderCaption = "Veckor";
        pt.GrandTotalCaption = "Totalt";

        pt = ws.PivotTables.Add(ws.Cells["A60"], wsData.Cells["K1:O11"], "Pivottable11");
        pt.RowFields.Add(pt.Fields["Category"]);
        pt.RowFields.Add(pt.Fields["Item"]);
        pt.RowFields.Add(pt.Fields[4]);

        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);

        pt.DataOnRows = true;
    }

    [TestMethod]
    public void Pivot_GroupNumber()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pivot-Group Number");
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:N11"], "Pivottable9");
        pt.PageFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[3]);
        pt.RowFields[0].AddNumericGrouping(-3.3, 5.5, 4.0);
        pt.DataFields.Add(pt.Fields[2]);
        pt.RowFields[0].Items[0].Hidden = true;
        pt.RowFields[0].Items[1].Hidden = true;
        pt.RowFields[0].Items[2].Hidden = true;
        pt.RowFields[0].Items[3].Hidden = true;
        pt.DataOnRows = false;
        pt.TableStyle = TableStyles.Medium14;
    }

    [TestMethod]
    public void Pivot_ManyRowFields()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pivot-Many RowFields");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable10");
        pt.ColumnFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        pt.RowFields.Add(pt.Fields[3]);
        pt.RowFields.Add(pt.Fields[2]);
        pt.RowFields.Add(pt.Fields[4]);
        pt.DataOnRows = true;
        pt.ColumnHeaderCaption = "Column Caption";
        pt.RowHeaderCaption = "Row Caption";
    }

    [TestMethod]
    public void Pivot_Blank()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pivot-Blank");

        wsData.Cells["A1"].Value = "Column1";
        wsData.Cells["B1"].Value = "Column2";
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["A1:B2"], "Pivottable11");
        pt.ColumnFields.Add(pt.Fields[1]);
        ExcelPivotTableField? rf = pt.RowFields.Add(pt.Fields[0]);
        rf.SubTotalFunctions = eSubTotalFunctions.None;
        pt.DataOnRows = true;
    }

    [TestMethod]
    public void Pivot_SaveDataFalse()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pivot-NoRecord");

        wsData.Cells["A1"].Value = "Column1";
        wsData.Cells["B1"].Value = "Column2";
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["A1:B2"], "Pivottable11");
        pt.ColumnFields.Add(pt.Fields[1]);
        ExcelPivotTableField? rf = pt.RowFields.Add(pt.Fields[0]);
        rf.SubTotalFunctions = eSubTotalFunctions.None;
        pt.DataOnRows = true;
        pt.CacheDefinition.SaveData = false;
    }

    [TestMethod]
    public void Pivot_SavedDataTrue()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pivot-WithRecord");

        wsData.Cells["A1"].Value = "Column1";
        wsData.Cells["B1"].Value = "Column2";
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["A1:B2"], "Pivottable11");
        pt.ColumnFields.Add(pt.Fields[1]);
        ExcelPivotTableField? rf = pt.RowFields.Add(pt.Fields[0]);
        rf.SubTotalFunctions = eSubTotalFunctions.None;
        pt.DataOnRows = true;
        pt.CacheDefinition.SaveData = false; //Remove the record xml
        pt.CacheDefinition.SaveData = true; //Add the record xml
    }

    [TestMethod]
    public void Pivot_ManyPageFields()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pivot-Many PageFields");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:O11"], "Pivottable12");
        pt.ColumnFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        ExcelPivotTableField? pf1 = pt.PageFields.Add(pt.Fields[2]);
        pf1.Items.Refresh();
        pf1.Items[1].Hidden = true;
        pf1.Items[8].Hidden = true;

        ExcelPivotTableField? pf2 = pt.PageFields.Add(pt.Fields[4]);
        pf2.Items.Refresh();
        pf2.Items[1].Hidden = true;
        pf1.MultipleItemSelectionAllowed = true;
        pf2.MultipleItemSelectionAllowed = true;
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataOnRows = true;
        pt.ColumnHeaderCaption = "Column Caption";
        pt.RowHeaderCaption = "Row Caption";

        Assert.AreEqual(1, pt.ColumnFields.Count);
        Assert.AreEqual(2, pt.PageFields.Count);
        Assert.AreEqual(1, pt.RowFields.Count);
        Assert.AreEqual(1, pt.DataFields.Count);
        Assert.IsTrue(pf1.MultipleItemSelectionAllowed);
    }

    [TestMethod]
    public void Pivot_StylingFieldsFalse()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pivot-StylingFieldsFalse");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:O11"], "Pivottable12");
        pt.ColumnFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        ExcelPivotTableDataField? df = pt.DataFields.Add(pt.Fields[3]);
        pt.DataOnRows = true;
        pt.ColumnHeaderCaption = "Column Caption";
        pt.RowHeaderCaption = "Row Caption";

        Assert.IsTrue(pt.ShowColumnHeaders);
        Assert.IsFalse(pt.ShowColumnStripes);
        Assert.IsTrue(pt.ShowRowHeaders);
        Assert.IsFalse(pt.ShowRowStripes);
        Assert.IsTrue(pt.ShowLastColumn);

        pt.ShowColumnHeaders = false;
        pt.ShowColumnStripes = true;
        pt.ShowRowHeaders = false;
        pt.ShowRowStripes = true;
        pt.ShowLastColumn = false;

        Assert.IsFalse(pt.ShowColumnHeaders);
        Assert.IsTrue(pt.ShowColumnStripes);
        Assert.IsFalse(pt.ShowRowHeaders);
        Assert.IsTrue(pt.ShowRowStripes);
        Assert.IsFalse(pt.ShowLastColumn);

        Assert.AreEqual(1, pt.ColumnFields.Count);
        Assert.AreEqual(1, pt.RowFields.Count);
        Assert.AreEqual(1, pt.DataFields.Count);
    }

    [TestMethod]
    public void RowsDataOnRow_WithNumberFormat()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotTable with numberformat");
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable2");
        pt.RowFields.Add(pt.Fields[1]);
        pt.RowFields.Add(pt.Fields[0]);
        pt.DataFields.Add(pt.Fields[3]);
        pt.DataFields.Add(pt.Fields[2]);

        pt.Fields[3].Format = "#,##0";
        pt.Fields[3].Cache.Format = "#,##0.000";
        ws.Workbook.Styles.UpdateXml();

        Assert.AreEqual(3, pt.Fields[3].NumFmtId);
        Assert.AreEqual(165, pt.Fields[3].Cache.NumFmtId);
    }

    [TestMethod]
    public void AddCalculatedField()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CalculatedField");

        LoadTestdata(ws);
        string? formula = "NumValue*2";
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        ExcelPivotTableField? f = tbl.Fields.AddCalculatedField("NumValueX2", formula);
        f.Format = "#,##0";
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df1 = tbl.DataFields.Add(tbl.Fields[3]);
        ExcelPivotTableDataField? df2 = tbl.DataFields.Add(tbl.Fields["NumValueX2"]);
        df1.Function = DataFieldFunctions.Sum;
        df2.Function = DataFieldFunctions.Sum;
        tbl.DataOnRows = false;
        Assert.AreEqual("NumValue*2", tbl.Fields[4].Cache.Formula);
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void ShouldThrowExceptionOnAddingCalculatedFieldToColumns()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("RowArgExcep");
        LoadTestdata(ws);
        string? formula = "NumValue*2";
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        tbl.Fields.AddCalculatedField("NumValueX2", formula);
        ExcelPivotTableField? rf = tbl.ColumnFields.Add(tbl.Fields[4]);
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void ShouldThrowExceptionOnAddingCalculatedFieldToRow()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("RowArgExcep");
        LoadTestdata(ws);
        string? formula = "NumValue*2";
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        tbl.Fields.AddCalculatedField("NumValueX2", formula);
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[4]);
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void ShouldThrowExceptionOnAddingCalculatedFieldToPage()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("RowArgExcep");
        LoadTestdata(ws);
        string? formula = "NumValue*2";
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        tbl.Fields.AddCalculatedField("NumValueX2", formula);
        ExcelPivotTableField? rf = tbl.PageFields.Add(tbl.Fields[4]);
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void ShouldThrowExceptionOnAddingCalculatedFieldWithBlankFormula()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("RowArgExcep");
        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        tbl.Fields.AddCalculatedField("NumValueX2", "");
    }

    [TestMethod]
    public void PivotTableStyleTests()
    {
        ExcelWorksheet? wsData = _pck.Workbook.Worksheets["Data"];
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleTests");

        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:N11"], "Pivottable8");
        pt.PivotTableStyle = PivotTableStyles.None;
        Assert.AreEqual(PivotTableStyles.None, pt.PivotTableStyle);
        Assert.AreEqual(TableStyles.None, pt.TableStyle);

        pt.PivotTableStyle = PivotTableStyles.Medium28;
        Assert.AreEqual(PivotTableStyles.Medium28, pt.PivotTableStyle);
        Assert.AreEqual(TableStyles.Medium28, pt.TableStyle);

        pt.PivotTableStyle = PivotTableStyles.Dark28;
        Assert.AreEqual(PivotTableStyles.Dark28, pt.PivotTableStyle);
        Assert.AreEqual(TableStyles.Custom, pt.TableStyle);
        Assert.AreEqual("PivotStyleDark28", pt.StyleName);

        pt.TableStyle = TableStyles.Light15;
        Assert.AreEqual(PivotTableStyles.Light15, pt.PivotTableStyle);
        Assert.AreEqual(TableStyles.Light15, pt.TableStyle);
        Assert.AreEqual("PivotStyleLight15", pt.StyleName);

        pt.PivotTableStyle = PivotTableStyles.Light28;
        Assert.AreEqual(PivotTableStyles.Light28, pt.PivotTableStyle);
        Assert.AreEqual(TableStyles.Custom, pt.TableStyle);
        Assert.AreEqual("PivotStyleLight28", pt.StyleName);
    }

    [TestMethod]
    public void HideValuesRow()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("HideValuesRow");

        LoadTestdata(ws);
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[1]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[3]);
        df.Function = DataFieldFunctions.Sum;
        tbl.DataOnRows = false;
        tbl.GridDropZones = false;
        Assert.IsTrue(tbl.ShowValuesRow);
        tbl.ShowValuesRow = false;
        Assert.IsFalse(tbl.ShowValuesRow);
        tbl.ShowValuesRow = true;
        Assert.IsTrue(tbl.ShowValuesRow);
        tbl.ShowValuesRow = false;
    }

    [TestMethod]
    public void ValidateSharedItemsAreCaseInsensitive()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("CaseInsentitive");

        ws.Cells["A1"].Value = "Column1";
        ws.Cells["B1"].Value = "Column2";
        ws.Cells["A2"].Value = "Value1";
        ws.Cells["B2"].Value = 1;
        ws.Cells["A3"].Value = "value1";
        ws.Cells["B3"].Value = 2;
        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:B3"], "CIPivotTable");
        ExcelPivotTableField? rf = tbl.RowFields.Add(tbl.Fields[0]);
        ExcelPivotTableDataField? df = tbl.DataFields.Add(tbl.Fields[1]);
        rf.Cache.Refresh();
        Assert.AreEqual(1, rf.Cache.SharedItems.Count);
        Assert.AreEqual("Value1", rf.Cache.SharedItems[0]);
    }

    [TestMethod]
    public void ValidateAttributesWhenNumbericAndMissing()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("NumericAndNull");
        ws.Cells["A1"].Value = "Int";
        ws.Cells["A2"].Value = 1;
        ws.Cells["A3"].Value = 2;
        ws.Cells["A4"].Value = 2;

        ws.Cells["B1"].Value = "Float";
        ws.Cells["B2"].Value = 1.3;
        ws.Cells["B3"].Value = 2.4;
        ws.Cells["B4"].Value = 5.6;

        ws.Cells["C1"].Value = "IntFloat";
        ws.Cells["C2"].Value = 3;
        ws.Cells["C3"].Value = 2.4;
        ws.Cells["C4"].Value = 2;

        ws.Cells["D1"].Value = "IntNull";
        ws.Cells["D2"].Value = 3;
        ws.Cells["D4"].Value = 3;

        ws.Cells["E1"].Value = "FloatNull";
        ws.Cells["E3"].Value = 4.2;
        ws.Cells["E4"].Value = 5.7;

        ws.Cells["F1"].Value = "IntFloatNull";
        ws.Cells["F2"].Value = 5;
        ws.Cells["F4"].Value = 6.2;

        ws.Cells["G1"].Value = "StringNull";
        ws.Cells["G2"].Value = "Value 1";
        ws.Cells["G4"].Value = "Value 3";

        ws.Cells["H1"].Value = "MixedIntBool";
        ws.Cells["H2"].Value = 1;
        ws.Cells["H4"].Value = true;

        ws.Cells["I1"].Value = "Mixed float";
        ws.Cells["I3"].Value = 3.3;
        ws.Cells["I4"].Value = "Value 3";

        ExcelPivotTable? tbl = ws.PivotTables.Add(ws.Cells["K3"], ws.Cells["A1:I4"], "ptNumberMissing");
        ExcelPivotTableField? pf1 = tbl.PageFields.Add(tbl.Fields[0]);
        ExcelPivotTableField? pf2 = tbl.PageFields.Add(tbl.Fields[1]);
        ExcelPivotTableField? pf3 = tbl.PageFields.Add(tbl.Fields[2]);
        ExcelPivotTableField? pf4 = tbl.PageFields.Add(tbl.Fields[3]);
        ExcelPivotTableField? pf5 = tbl.PageFields.Add(tbl.Fields[4]);
        ExcelPivotTableField? pf6 = tbl.PageFields.Add(tbl.Fields[5]);
        ExcelPivotTableField? pf7 = tbl.PageFields.Add(tbl.Fields[6]);
        ExcelPivotTableField? pf8 = tbl.PageFields.Add(tbl.Fields[7]);
        ExcelPivotTableField? pf9 = tbl.PageFields.Add(tbl.Fields[8]);

        tbl.CacheDefinition.Refresh();

        p.Save();

        AssertShartedItemsAttributes(pf1.Cache.TopNode.FirstChild, 4, true, true, false, false, false);
        AssertShartedItemsAttributes(pf2.Cache.TopNode.FirstChild, 3, true, false, false, false, false);
        AssertShartedItemsAttributes(pf3.Cache.TopNode.FirstChild, 3, true, false, false, false, false);
        AssertShartedItemsAttributes(pf4.Cache.TopNode.FirstChild, 4, true, true, true, false, false);
        AssertShartedItemsAttributes(pf5.Cache.TopNode.FirstChild, 3, true, false, true, false, false);
        AssertShartedItemsAttributes(pf6.Cache.TopNode.FirstChild, 3, true, false, true, false, false);
        AssertShartedItemsAttributes(pf7.Cache.TopNode.FirstChild, 1, false, false, true, false, false);

        AssertShartedItemsAttributes(pf8.Cache.TopNode.FirstChild, 4, true, true, true, false, true);
        AssertShartedItemsAttributes(pf9.Cache.TopNode.FirstChild, 3, true, false, true, false, true);
    }

    private static void AssertShartedItemsAttributes(XmlNode node,
                                                     int count,
                                                     bool numberValues,
                                                     bool intValues,
                                                     bool containsBlanks,
                                                     bool semiMixedValues,
                                                     bool mixedValues)
    {
        if (node.Attributes.Count != count)
        {
            Assert.Fail("Wrong attrib Count");
        }

        AssertContains(node, "containsNumber", numberValues);
        AssertContains(node, "containsInteger", intValues);
        AssertContains(node, "containsBlank", containsBlanks);
        AssertContains(node, "containsSemiMixedTypes", semiMixedValues);
        AssertContains(node, "containsMixedTypes", mixedValues);

        //containsInteger = "1" containsNumber = "1" containsString = "0" containsSemiMixedTypes = "0"
    }

    private static void AssertContains(XmlNode node, string attrName, bool value)
    {
        XmlAttribute? a = node.Attributes[attrName];

        if (a == null)
        {
            if (value)
            {
                Assert.Fail($"{attrName} value not false");
            }
        }
        else
        {
            if (value && a.Value != "1")
            {
                Assert.Fail($"{attrName} value not true");
            }
        }
    }

    [TestMethod]
    public void CopyPivotTableToExternalPackageSameWorksheetAsData()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("data");
        LoadTestdata(ws);

        ExcelPivotTable? pivotTable1 = ws.PivotTables.Add(ws.Cells["G1"], ws.Cells["A1:D100"], "PivotTable1");
        pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
        pivotTable1.RowFields.Add(pivotTable1.Fields[2]);
        pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
        pivotTable1.DataFields.Add(pivotTable1.Fields[3]);

        using (ExcelPackage? p2 = new ExcelPackage())
        {
            ExcelWorksheet? wsNew = p2.Workbook.Worksheets.Add("PivotCopy", ws);
            SaveWorkbook("copiedPivot.xlsx", p2);
        }

        SaveWorkbook("Pivot.xlsx", p);
    }
}