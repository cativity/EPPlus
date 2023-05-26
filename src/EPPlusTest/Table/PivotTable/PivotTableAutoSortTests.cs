using FakeItEasy.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableAutoSortTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableAutoSort.xlsx", true);
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
        public void SetAutoSortAcending()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortAcending");
            ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot1");
            ExcelPivotTableField? rf=p1.RowFields.Add(p1.Fields[0]);
            ExcelPivotTableDataField? df=p1.DataFields.Add(p1.Fields[3]);
            rf.SetAutoSort(df);
        }
        [TestMethod]
        public void SetAutoSortDesending()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescending");
            ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            ExcelPivotTableField? rf = p1.RowFields.Add(p1.Fields[0]);
            ExcelPivotTableDataField? df = p1.DataFields.Add(p1.Fields[3]);
            rf.SetAutoSort(df, eSortType.Descending);
        }
        [TestMethod]
        public void SetAutoSortDataAndColumnField1()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescendingCF1");
            ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");            
            ExcelPivotTableField? rf = p1.RowFields.Add(p1.Fields[0]);            
            ExcelPivotTableField? cf = p1.ColumnFields.Add(p1.Fields[1]);
            ExcelPivotTableDataField? df = p1.DataFields.Add(p1.Fields[3]);
            rf.SetAutoSort(df, eSortType.Descending);
            ExcelPivotAreaReference? reference = rf.AutoSort.Conditions.Fields.Add(cf);
            cf.Items.Refresh();
            reference.Items.AddByValue("Hardware");
        }
        [TestMethod]
        public void SetAutoSortDataAndColumnField2()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescendingCF2");
            ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            ExcelPivotTableField? rf = p1.RowFields.Add(p1.Fields[0]);
            ExcelPivotTableField? cf = p1.ColumnFields.Add(p1.Fields[1]);
            ExcelPivotTableDataField? df = p1.DataFields.Add(p1.Fields[3]);
            rf.SetAutoSort(df, eSortType.Descending);
            ExcelPivotAreaReference? reference = rf.AutoSort.Conditions.Fields.Add(cf);
            cf.Items.Refresh();
            reference.Items.Add(1);
        }
        [TestMethod]
        public void SetAutoSortDataAndRowField1()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescendingRF1");
            ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            ExcelPivotTableField? rf = p1.RowFields.Add(p1.Fields[0]);
            ExcelPivotTableField? cf = p1.ColumnFields.Add(p1.Fields[1]);
            ExcelPivotTableDataField? df = p1.DataFields.Add(p1.Fields[3]);
            cf.SetAutoSort(df, eSortType.Descending);
            ExcelPivotAreaReference? reference = cf.AutoSort.Conditions.Fields.Add(rf);
            rf.Items.Refresh();
            reference.Items.Add(0);
        }
        [TestMethod]
        public void SetAutoSortDataAndRowField3()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescendingRF3");
            ExcelPivotTable? p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            ExcelPivotTableField? rf = p1.RowFields.Add(p1.Fields[0]);
            ExcelPivotTableField? cf = p1.ColumnFields.Add(p1.Fields[1]);
            ExcelPivotTableDataField? df = p1.DataFields.Add(p1.Fields[3]);
            cf.SetAutoSort(df, eSortType.Descending);
            ExcelPivotAreaReference? reference = cf.AutoSort.Conditions.Fields.Add(rf);
            rf.Items.Refresh();
            reference.Items.Add(2);
        }
        [TestMethod]
        public void ReadAutoSort()
        {
            using ExcelPackage? p1 = new ExcelPackage();
            ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("PivotSameAutoClear");
            ExcelRangeBase? r = LoadItemData(ws);
            ws.Tables.Add(r, "Table1");

            ExcelPivotTable? pivot1 = ws.PivotTables.Add(ws.Cells["A1"], p1.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            ExcelPivotTableField? rf = pivot1.RowFields.Add(pivot1.Fields[0]);
            ExcelPivotTableField? cf = pivot1.ColumnFields.Add(pivot1.Fields[1]);
            ExcelPivotTableDataField? df = pivot1.DataFields.Add(pivot1.Fields[3]);
            cf.SetAutoSort(df, eSortType.Descending);
            ExcelPivotAreaReference? reference = cf.AutoSort.Conditions.Fields.Add(rf);
            rf.Items.Refresh();
            reference.Items.Add(2);

            Assert.IsNotNull(cf.AutoSort);

            p1.Save();

            using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
            ExcelWorksheet? ws1 = p1.Workbook.Worksheets[0];
            ExcelPivotTable? pivot2 = ws.PivotTables[0];

            Assert.AreEqual(1, pivot2.ColumnFields.Count);
            Assert.AreEqual(1, pivot2.RowFields.Count);
            Assert.AreEqual(1, pivot2.DataFields.Count);
            Assert.IsNotNull(pivot2.ColumnFields[0].AutoSort);
            Assert.AreEqual(1, pivot2.ColumnFields[0].AutoSort.Conditions.DataFields.Count);
            Assert.AreEqual(1, pivot2.ColumnFields[0].AutoSort.Conditions.Fields.Count);
        }
        [TestMethod]
        public void RemoveAutoSort()
        {
            using ExcelPackage? p1 = new ExcelPackage();
            ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("PivotSameAutoClear");
            ExcelRangeBase? r = LoadItemData(ws);
            ws.Tables.Add(r, "Table1");

            ExcelPivotTable? pivot1 = ws.PivotTables.Add(ws.Cells["A1"], p1.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            ExcelPivotTableField? rf = pivot1.RowFields.Add(pivot1.Fields[0]);
            ExcelPivotTableField? cf = pivot1.ColumnFields.Add(pivot1.Fields[1]);
            ExcelPivotTableDataField? df = pivot1.DataFields.Add(pivot1.Fields[3]);
            cf.SetAutoSort(df, eSortType.Descending);
            ExcelPivotAreaReference? reference = cf.AutoSort.Conditions.Fields.Add(rf);
            rf.Items.Refresh();
            reference.Items.Add(2);

            Assert.IsNotNull(cf.AutoSort);

            p1.Save();

            using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
            ExcelWorksheet? ws1 = p1.Workbook.Worksheets[0];
            ExcelPivotTable? pivot2 = ws.PivotTables[0];

            Assert.AreEqual(1, pivot2.ColumnFields.Count);
            Assert.AreEqual(1, pivot2.RowFields.Count);
            Assert.AreEqual(1, pivot2.DataFields.Count);
            Assert.IsNotNull(pivot2.ColumnFields[0].AutoSort);

            pivot2.ColumnFields[0].RemoveAutoSort();
            Assert.IsNull(pivot2.ColumnFields[0].AutoSort);
        }
    }
}
