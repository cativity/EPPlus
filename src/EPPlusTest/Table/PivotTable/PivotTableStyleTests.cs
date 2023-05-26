using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableStyleTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableStyle.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Data1");
            LoadItemData(_ws);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            string? dirName = _pck.File.DirectoryName;
            string? fileName = _pck.File.FullName;
            SaveAndCleanup(_pck);
            File.Copy(fileName, dirName + "\\PivotTableReadStyle.xlsx", true);
        }
        internal ExcelPivotTable CreatePivotTable(ExcelWorksheet ws)
        {
            ExcelPivotTable? pt = ws.PivotTables.Add(ws.Cells["A3"], _ws.Cells[_ws.Dimension.Address], "PivotTable1");
            pt.RowFields.Add(pt.Fields[0]);
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.PageFields.Add(pt.Fields[4]);
            return pt;
        }
        [TestMethod]
        public void AddPivotAllStyle()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleAll");
            ExcelPivotTable? pt = CreatePivotTable(ws);
            ExcelPivotTableAreaStyle? s=pt.Styles.AddWholeTable();
            s.Style.Font.Name = "Bauhaus 93";

            Assert.IsTrue(s.Style.HasValue);
            Assert.AreEqual("Bauhaus 93", s.Style.Font.Name);
        }
        [TestMethod]
        public void AddPivotLabels()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleAllLabels");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddAllLabels();
            s.Style.Font.Color.SetColor(Color.Green);
        }
        [TestMethod]
        public void AddPivotAllData()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleAllData");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddAllData();
            s.Style.Font.Color.SetColor(Color.Blue);
        }

        [TestMethod]
        public void AddPivotLabelPageField()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StylePageFieldLabel");
            ExcelPivotTable? pt = CreatePivotTable(ws);
            ExcelPivotTableAreaStyle? s = pt.Styles.AddLabel(pt.PageFields[0]);
            s.Style.Font.Color.SetColor(Color.Green);
        }
        [TestMethod]
        public void AddPivotLabelColumnField()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleColumnFieldLabel");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddLabel(pt.ColumnFields[0]);
            s.Style.Font.Color.SetColor(Color.Indigo);
        }
        [TestMethod]
        public void AddPivotLabelColumnFieldSingleCell()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleColumnFieldLabelCell");
            ExcelPivotTable? pt = CreatePivotTable(ws);
            pt.DataOnRows = false;
            pt.CacheDefinition.Refresh();
            ExcelPivotTableAreaStyle? s = pt.Styles.AddLabel(pt.ColumnFields[0]);
            s.Conditions.DataFields.Add(0);
            s.Conditions.Fields[0].Items.Add(0);
            s.Conditions.Fields[0].Items.Add(1);
            s.Style.Font.Color.SetColor(Color.Indigo);
        }

        [TestMethod]
        public void AddPivotLabelRowColumnField()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleRowFieldLabel");
            ExcelPivotTable? pt = CreatePivotTable(ws);
            ExcelPivotTableAreaStyle? s = pt.Styles.AddLabel(pt.RowFields[0]);

            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotDataRowColumnField()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleRowFieldData");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddData(pt.RowFields[0]);

            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotData()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleData");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddData(pt.Fields[0], pt.Fields[1]);
            s.Style.Fill.Style = eDxfFillStyle.PatternFill;
            s.Style.Fill.BackgroundColor.SetColor(Color.Red);
            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotDataGrandColumn()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleDataGrandColumn");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddData(pt.Fields[0], pt.Fields[1]);
            s.GrandColumn = true;
            s.Style.Fill.Style = eDxfFillStyle.PatternFill;
            s.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            s.Style.Font.Underline = ExcelUnderLineType.Single;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotDataGrandRow()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleDataGrandRow");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddData();
            s.GrandRow = true;
            s.Style.Fill.Style = eDxfFillStyle.PatternFill;
            s.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            s.Style.Font.Underline = ExcelUnderLineType.Single;
            s.Style.Font.Name = "Times New Roman";
        }

        [TestMethod]
        public void AddPivotLabelRow()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleRowFieldLabelTot");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddLabel(pt.RowFields[0]);
            s.GrandRow = true;
            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotLabelRowDf1()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleRowFieldLabelTotDf1");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddLabel(pt.RowFields[0]);
            s.Conditions.DataFields.Add(1);
            s.GrandRow = true;
            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }

        [TestMethod]
        public void AddPivotLabelRowDataField2()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleRowFieldDf2");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddLabel(pt.RowFields[0]);
            s.Conditions.DataFields.Add(1);
            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotLabelRowDataField2AndValue()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleRowFieldDf2Value");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            pt.CacheDefinition.Refresh();
            ExcelPivotTableAreaStyle? s = pt.Styles.AddLabel(pt.RowFields[0]);
            s.Conditions.DataFields.Add(1);
            s.Conditions.Fields[0].Items.AddByValue("Screwdriver");
            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotDataItemByIndex()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotDataItemIndex");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            pt.CacheDefinition.Refresh();
            ExcelPivotTableAreaStyle? s = pt.Styles.AddData(pt.Fields[0], pt.Fields[1]);
            s.Conditions.DataFields.Add(0);
            s.Conditions.Fields[0].Items.Add(0);
            s.Conditions.Fields[1].Items.Add(0);
            s.Style.Fill.Style = eDxfFillStyle.PatternFill;
            s.Style.Fill.BackgroundColor.SetColor(Color.Red);
            s.Outline = true;
            //s.Axis = ePivotTableAxis.RowAxis;
            s.Style.Font.Color.SetColor(Color.Blue);
        }
        [TestMethod]
        public void AddPivotDataItemByValue()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("PivotDataItemValue");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            pt.CacheDefinition.Refresh();
            ExcelPivotTableAreaStyle? s = pt.Styles.AddData(pt.Fields[0], pt.Fields[1]);
            s.Conditions.DataFields.Add(pt.DataFields[1]);
            s.Conditions.Fields[0].Items.AddByValue("Apple");
            s.Conditions.Fields[1].Items.AddByValue("Groceries");
            s.Style.Fill.Style = eDxfFillStyle.PatternFill;
            s.Style.Fill.BackgroundColor.SetColor(Color.Red);
            s.Outline = true;
            //s.Axis = ePivotTableAxis.RowAxis;
            s.Style.Font.Color.SetColor(Color.Blue);
        }

        [TestMethod]
        public void AddButtonField()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleFieldPage");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddButtonField(pt.Fields[4]);
            s.Style.Font.Color.SetColor(Color.Pink);
        }

        [TestMethod]
        public void AddButtonRowAxis()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleButtonRowAxis");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s = pt.Styles.AddButtonField(ePivotTableAxis.RowAxis);
            s.Style.Font.Underline = ExcelUnderLineType.DoubleAccounting;
        }
        [TestMethod]
        public void AddButtonColumnAxis()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleButtonColumnAxis");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s3 = pt.Styles.AddButtonField(ePivotTableAxis.ColumnAxis);
            s3.Style.Font.Italic = true;
        }
        [TestMethod]
        public void AddButtonPageAxis()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleButtonPageAxis");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? s4 = pt.Styles.AddButtonField(ePivotTableAxis.PageAxis);
            s4.Style.Font.Color.SetColor(Color.ForestGreen);
        }


        [TestMethod]
        public void AddTopStart()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleTopStart");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            //Top Left cells 
            ExcelPivotTableAreaStyle? styleTopLeft = pt.Styles.AddTopStart();
            styleTopLeft.Style.Fill.PatternType = ExcelFillStyle.Solid;
            styleTopLeft.Style.Fill.BackgroundColor.SetColor(Color.Red);
        }
        [TestMethod]
        public void AddTopStartOffset0()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleTopStartOffset0");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            //Top Left cells 
            ExcelPivotTableAreaStyle? styleTopLeft = pt.Styles.AddTopStart("A1");
            styleTopLeft.Style.Fill.PatternType = ExcelFillStyle.Solid;
            styleTopLeft.Style.Fill.BackgroundColor.SetColor(Color.Blue);
        }

        [TestMethod]
        public void AddTopEnd()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleTopEnd");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? styleTopRight2 = pt.Styles.AddTopEnd();
            styleTopRight2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            styleTopRight2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        }
        [TestMethod]
        public void AddTopEndOffset1()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("StyleTopEndOffset1");
            ExcelPivotTable? pt = CreatePivotTable(ws);

            ExcelPivotTableAreaStyle? styleTopRight2 = pt.Styles.AddTopEnd("A1");
            styleTopRight2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            styleTopRight2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);            
        }
    }
}

