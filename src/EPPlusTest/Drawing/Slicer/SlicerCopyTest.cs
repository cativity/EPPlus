using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Drawing.Slicer;

[TestClass]
public class SlicerCopyTest : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        _pck = OpenPackage("SlicerCopy.xlsx", true);
    }

    [ClassCleanup]
    public static void Cleanup()
    {
        SaveAndCleanup(_pck);
    }

    [TestMethod]
    public void CopyTableSlicer()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("TableSlicerSource");

        LoadTestdata(ws);
        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table2");
        ExcelTableSlicer? slicer = ws.Drawings.AddTableSlicer(tbl.Columns[1]);
        slicer.SetPosition(1, 0, 5, 0);

        slicer.SetSize(200, 600);

        _ = _pck.Workbook.Worksheets.Add("TableSlicerCopy", ws);
    }

    [TestMethod]
    public void CopyPivotTableSlicer()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotTableSlicerSource");

        LoadTestdata(ws);
        ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Table3");
        _ = pt.RowFields.Add(pt.Fields[1]);
        _ = pt.DataFields.Add(pt.Fields[3]);
        ExcelPivotTableSlicer? slicer = ws.Drawings.AddPivotTableSlicer(pt.Fields[3]);
        slicer.SetPosition(1, 0, 8, 0);

        slicer.SetSize(200, 600);

        _ = _pck.Workbook.Worksheets.Add("PivotTableSlicerCopy", ws);
    }
}