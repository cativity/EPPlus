using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableStyleReadTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableReadStyle.xlsx");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestMethod]
        public void ReadPivotAllStyle()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleAll");
            ExcelPivotTable? pt = ws.PivotTables[0];
            Assert.AreEqual(1, pt.Styles.Count);

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.All, s.PivotAreaType);
            Assert.IsTrue(s.Style.HasValue);
            Assert.AreEqual("Bauhaus 93", s.Style.Font.Name);
        }
        [TestMethod]
        public void ReadPivotLabels()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleAllLabels");
            ExcelPivotTable? pt = ws.PivotTables[0];
            Assert.AreEqual(1, pt.Styles.Count);

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.All, s.PivotAreaType);
            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);

            Assert.AreEqual(Color.Green.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }
        [TestMethod]
        public void ReadPivotAllData()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleAllData");
            ExcelPivotTable? pt = ws.PivotTables[0];
            Assert.AreEqual(1, pt.Styles.Count);

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.All, s.PivotAreaType);
            Assert.IsTrue(s.DataOnly);
            Assert.IsFalse(s.LabelOnly);

            Assert.AreEqual(Color.Blue.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }

        [TestMethod]
        public void ReadPivotLabelPageField()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StylePageFieldLabel");
            ExcelPivotTable? pt = ws.PivotTables[0];
            Assert.AreEqual(1, pt.Styles.Count);
            Assert.AreEqual(1, pt.Styles.Count);

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.Normal, s.PivotAreaType);
            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.PageFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.AreEqual(Color.Green.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }
        [TestMethod]
        public void ReadPivotLabelColumnField()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleColumnFieldLabel");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.ColumnFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.AreEqual(Color.Indigo.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }
        [TestMethod]
        public void AddPivotLabelColumnFieldSingleCell()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleColumnFieldLabelCell");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.IsFalse(pt.DataOnRows);
            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual("Price", s.Conditions.DataFields[0].Name);

            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(2, s.Conditions.Fields[0].Items.Count);
            Assert.AreEqual(0, s.Conditions.Fields[0].Items[0].Index);
            Assert.AreEqual(1, s.Conditions.Fields[0].Items[1].Index);

            Assert.AreEqual(Color.Indigo.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }

        [TestMethod]
        public void AddPivotLabelRowColumnField()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleRowFieldLabel");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.Normal, s.PivotAreaType);
            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void ReadPivotDataRowColumnField()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleRowFieldData");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.Data, s.PivotAreaType);
            Assert.IsTrue(s.DataOnly);
            Assert.IsFalse(s.LabelOnly);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void ReadPivotData()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleData");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(2, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.Fields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual(pt.Fields[1].Index, s.Conditions.Fields[1].FieldIndex);

            Assert.AreEqual(s.Style.Fill.Style, eDxfFillStyle.PatternFill);
            Assert.AreEqual(Color.Red.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());
            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void AddPivotDataGrandColumn()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleDataGrandColumn");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(2, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.Fields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual(pt.Fields[1].Index, s.Conditions.Fields[1].FieldIndex);
            Assert.IsTrue(s.GrandColumn);
            Assert.AreEqual(s.Style.Fill.Style, OfficeOpenXml.Style.eDxfFillStyle.PatternFill);
            Assert.AreEqual(Color.LightGray.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());

            Assert.AreEqual(ExcelUnderLineType.Single, s.Style.Font.Underline);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void AddPivotDataGrandRow()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleDataGrandRow");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.IsTrue(s.GrandRow);
            Assert.AreEqual(s.Style.Fill.Style, OfficeOpenXml.Style.eDxfFillStyle.PatternFill);
            Assert.AreEqual(Color.LightGray.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());

            Assert.AreEqual(ExcelUnderLineType.Single, s.Style.Font.Underline);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }

        [TestMethod]
        public void AddPivotLabelRow()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleRowFieldLabelTot");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.IsTrue(s.GrandRow);
            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void AddPivotLabelRowDf1()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleRowFieldLabelTotDf1");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Name);

            Assert.IsTrue(s.GrandRow);
            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }

        [TestMethod]
        public void ReadPivotLabelRowDataField2()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleRowFieldDf2");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Name);

            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void ReadPivotLabelRowDataField2AndValue()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleRowFieldDf2Value");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(1, s.Conditions.Fields[0].Items.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Name);

            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Name);

            Assert.AreEqual("Screwdriver", s.Conditions.Fields[0].Items[0].Value);

            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);

        }
        [TestMethod]
        public void ReadPivotDataItemByIndex()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "PivotDataItemIndex");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(2, s.Conditions.Fields.Count);
            Assert.AreEqual(1, s.Conditions.Fields[0].Items.Count);
            Assert.AreEqual(1, s.Conditions.Fields[1].Items.Count);
            Assert.AreEqual(0, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual(1, s.Conditions.Fields[1].FieldIndex);

            Assert.AreEqual(0, s.Conditions.Fields[0].Items[0].Index);
            Assert.AreEqual(0, s.Conditions.Fields[1].Items[0].Index);

            Assert.AreEqual(eDxfFillStyle.PatternFill, s.Style.Fill.Style);
            Assert.AreEqual(Color.Red.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());            
            Assert.IsTrue(s.Outline);
            Assert.AreEqual(Color.Blue.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());            
        }
        [TestMethod]
        public void ReadPivotDataItemByValue()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "PivotDataItemValue");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(2, s.Conditions.Fields.Count);
            Assert.AreEqual(1, s.Conditions.Fields[0].Items.Count);
            Assert.AreEqual(1, s.Conditions.Fields[1].Items.Count);
            Assert.AreEqual("Apple", s.Conditions.Fields[0].Items[0].Value);
            Assert.AreEqual("Groceries", s.Conditions.Fields[1].Items[0].Value);
            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Field.Name);

            Assert.AreEqual(eDxfFillStyle.PatternFill, s.Style.Fill.Style);
            Assert.AreEqual(Color.Red.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());
            Assert.IsTrue(s.Outline);
            Assert.AreEqual(Color.Blue.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }

        [TestMethod]
        public void ReadFieldButton()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleFieldPage");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.FieldButton, s.PivotAreaType);
            Assert.AreEqual(0, s.Conditions.DataFields.Count);
            Assert.AreEqual(4, s.FieldIndex);

            Assert.AreEqual(Color.Pink.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }

        [TestMethod]
        public void ReadButtonRowAxis()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleButtonRowAxis");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.FieldButton, s.PivotAreaType);
            Assert.AreEqual(ePivotTableAxis.RowAxis, s.Axis);
            Assert.AreEqual(ExcelUnderLineType.DoubleAccounting, s.Style.Font.Underline);
        }
        [TestMethod]
        public void ReadButtonColumnAxis()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleButtonColumnAxis");
            ExcelPivotTable? pt = ws.PivotTables[0];

            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.FieldButton, s.PivotAreaType);
            Assert.AreEqual(ePivotTableAxis.ColumnAxis, s.Axis);
            
            Assert.IsTrue(s.Style.Font.Italic.Value);
        }
        [TestMethod]
        public void ReadButtonPageAxis()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleButtonPageAxis");
            ExcelPivotTable? pt = ws.PivotTables[0];
            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.FieldButton, s.PivotAreaType);
            Assert.AreEqual(ePivotTableAxis.PageAxis, s.Axis);

            Assert.AreEqual(Color.ForestGreen.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());

       }
        [TestMethod]
        public void ReadTopStart()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleTopStart");
            ExcelPivotTable? pt = ws.PivotTables[0];
            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            //Top Left cells 
            Assert.AreEqual(ePivotAreaType.Origin, s.PivotAreaType);
            Assert.AreEqual(eDxfFillStyle.PatternFill, s.Style.Fill.Style);
            Assert.AreEqual(Color.Red.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        }
        [TestMethod]
        public void ReadTopStartOffset0()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleTopStartOffset0");
            ExcelPivotTable? pt = ws.PivotTables[0];
            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            //Top Left cells
            Assert.AreEqual(ePivotAreaType.Origin, s.PivotAreaType);
            Assert.AreEqual("A1", s.Offset);
            Assert.AreEqual(eDxfFillStyle.PatternFill, s.Style.Fill.Style);
            Assert.AreEqual(Color.Blue.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        }

        [TestMethod]
        public void AddTopEnd()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleTopEnd");
            ExcelPivotTable? pt = ws.PivotTables[0];
            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.TopEnd, s.PivotAreaType);
            Assert.AreEqual(eDxfFillStyle.PatternFill, s.Style.Fill.Style);
            Assert.AreEqual(Color.Yellow.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        }
        [TestMethod]
        public void AddTopEndOffset1()
        {
            ExcelWorksheet? ws = TryGetWorksheet(_pck, "StyleTopEndOffset1");
            ExcelPivotTable? pt = ws.PivotTables[0];
            ExcelPivotTableAreaStyle? s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.TopEnd, s.PivotAreaType);
            Assert.AreEqual("A1", s.Offset);
            Assert.AreEqual(eDxfFillStyle.PatternFill, s.Style.Fill.Style);
            Assert.AreEqual(Color.Yellow.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());
        }

    }
}