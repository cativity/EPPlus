﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
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

namespace EPPlusTest.Table.PivotTable;

[TestClass]
public class PivotTableLargeStyleTests : TestBase
{
    [ClassInitialize]
    public static void Init(TestContext context) => InitBase();

    [ClassCleanup]
    public static void Cleanup()
    {
    }

    [TestMethod]
    public void AddPivotAllStyle()
    {
        using ExcelPackage? p = OpenTemplatePackage("PivotStyleLarge.xlsx");
        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        ExcelPivotTable? pt = ws.PivotTables[0];

        ExcelPivotTableAreaStyle? s0 = pt.Styles.AddButtonField(ePivotTableAxis.PageAxis, 2);
        s0.Style.Font.Color.SetColor(Color.Pink);

        ExcelPivotTableAreaStyle? s1 = pt.Styles.AddButtonField(pt.Fields["FacilityName"]);
        s1.Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Accent1);

        ExcelPivotTableAreaStyle? s2 = pt.Styles.AddLabel(pt.Fields["FacilityName"]);
        s2.Style.Font.Color.SetColor(Color.Green);

        ExcelPivotTableAreaStyle? s3 = pt.Styles.AddButtonField(pt.Fields["SiteId"]);
        s3.Style.Font.Color.SetColor(Color.Blue);

        ExcelPivotTableAreaStyle? s4 = pt.Styles.AddLabel(pt.Fields["SiteId"]);
        s4.Style.Font.Color.SetColor(Color.Cyan);
        _ = s4.Conditions.Fields[0].Items.AddByValue(5D);
        _ = s4.Conditions.Fields[0].Items.AddByValue(8D);
        _ = s4.Conditions.Fields[0].Items.AddByValue(9D);

        ExcelPivotTableAreaStyle? s5 = pt.Styles.AddData(pt.Fields["SiteId"], pt.Fields["ZipCode"], pt.Fields["Id"]);
        s5.Style.Fill.PatternType = ExcelFillStyle.DarkTrellis;
        s5.Style.Fill.BackgroundColor.SetColor(Color.Red);
        s5.Conditions.DataFields.Add(1);
        _ = s5.Conditions.Fields[0].Items.AddByValue(1D);
        _ = s5.Conditions.Fields[0].Items.AddByValue(2D);
        _ = s5.Conditions.Fields[0].Items.AddByValue(3D);
        _ = s5.Conditions.Fields[1].Items.AddByValue("02201");
        _ = s5.Conditions.Fields[2].Items.AddByValue("1100");

        ExcelPivotTableAreaStyle? s6 = pt.Styles.AddLabel(pt.Fields["ZipCode"], pt.Fields["Id"]);
        s6.Style.Fill.PatternType = ExcelFillStyle.LightUp;
        s6.Style.Fill.BackgroundColor.SetColor(Color.Green);

        //s6.Conditions.DataFields.Add(1);
        _ = s6.Conditions.Fields[0].Items.AddByValue("02201");
        s6.Offset = "B1:C1";
        SaveWorkbook("PivotStyleLargeSaved.xlsx", p);
    }
}