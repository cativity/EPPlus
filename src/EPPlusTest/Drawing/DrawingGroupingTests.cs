﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.VBA;
using System;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Table;

namespace EPPlusTest.Drawing.Grouping;

[TestClass]
public class DrawingGroupingTests : TestBase
{
    static ExcelPackage _pck;
    static ExcelWorksheet _ws;

    [ClassInitialize]
    public static void Init(TestContext context) => _pck = OpenPackage("DrawingGrouping.xlsx", true);

    [ClassCleanup]
    public static void Cleanup() => SaveAndCleanup(_pck);

    [TestMethod]
    public void Group_GroupBoxWithRadioButtonsTest()
    {
        _ws = _pck.Workbook.Worksheets.Add("GroupBox");
        ExcelControlGroupBox? ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
        ctrl.Text = "Groupbox 1";
        ctrl.SetPosition(480, 80);
        ctrl.SetSize(200, 120);

        _ws.Cells["G1"].Value = "Linked Groupbox";
        ctrl.LinkedCell = _ws.Cells["G1"];

        ExcelControlRadioButton? r1 = _ws.Drawings.AddRadioButtonControl("Option Button 1");
        r1.SetPosition(500, 100);
        r1.SetSize(100, 25);
        ExcelControlRadioButton? r2 = _ws.Drawings.AddRadioButtonControl("Option Button 2");
        r2.SetPosition(530, 100);
        r2.SetSize(100, 25);
        ExcelControlRadioButton? r3 = _ws.Drawings.AddRadioButtonControl("Option Button 3");
        r3.SetPosition(560, 100);
        r3.SetSize(100, 25);
        r3.FirstButton = true;

        _ = ctrl.Group(r1, r2, r3);
    }

    [TestMethod]
    public void Group_SingleDrawing()
    {
        _ws = _pck.Workbook.Worksheets.Add("SingleDrawing");
        ExcelControlGroupBox? ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
        ctrl.SetPosition(480, 80);
        ctrl.SetSize(200, 120);

        _ = ctrl.Group();
    }

    [TestMethod]
    public void Group_AddControlViaGroupShape()
    {
        _ws = _pck.Workbook.Worksheets.Add("AddViaGroupShape");
        ExcelControlGroupBox? ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
        ctrl.SetPosition(480, 80);
        ctrl.SetSize(200, 120);

        ExcelControlRadioButton? r1 = _ws.Drawings.AddRadioButtonControl("Option Button 1");
        r1.SetPosition(500, 100);
        r1.SetSize(100, 25);

        ExcelGroupShape? group = ctrl.Group();
        group.Drawings.Add(r1);
    }

    [TestMethod]
    public void UnGroup_SingleDrawing()
    {
        _ws = _pck.Workbook.Worksheets.Add("UnGroupSingleDrawing");
        ExcelControlGroupBox? ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
        ctrl.SetPosition(480, 80);
        ctrl.SetSize(200, 120);

        _ = ctrl.Group();
        ctrl.UnGroup();
    }

    [TestMethod]
    public void UnGroup_GroupBoxWithRadioButtonsTest()
    {
        _ws = _pck.Workbook.Worksheets.Add("UnGroupAllDrawings");
        ExcelControlGroupBox? ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
        ctrl.Text = "Groupbox 1";
        ctrl.SetPosition(480, 80);
        ctrl.SetSize(200, 120);

        _ws.Cells["G1"].Value = "Linked Groupbox";
        ctrl.LinkedCell = _ws.Cells["G1"];

        ExcelControlRadioButton? r1 = _ws.Drawings.AddRadioButtonControl("Option Button 1");
        r1.SetPosition(500, 100);
        r1.SetSize(100, 25);
        ExcelControlRadioButton? r2 = _ws.Drawings.AddRadioButtonControl("Option Button 2");
        r2.SetPosition(530, 100);
        r2.SetSize(100, 25);
        ExcelControlRadioButton? r3 = _ws.Drawings.AddRadioButtonControl("Option Button 3");
        r3.SetPosition(560, 100);
        r3.SetSize(100, 25);
        r3.FirstButton = true;

        ExcelGroupShape? g = ctrl.Group(r1, r2, r3);

        g.SetPosition(100, 100); //Move whole group

        r1.UnGroup(false);
    }

    [TestMethod]
    public void Group_GroupIntoGroupTest()
    {
        _ws = _pck.Workbook.Worksheets.Add("GroupIntoGroup");
        ExcelControlGroupBox? ctrl = (ExcelControlGroupBox)_ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
        ctrl.Text = "Groupbox 1";
        ctrl.SetPosition(480, 80);
        ctrl.SetSize(200, 120);

        _ws.Cells["G1"].Value = "Linked Groupbox";
        ctrl.LinkedCell = _ws.Cells["G1"];

        ExcelControlRadioButton? r1 = _ws.Drawings.AddRadioButtonControl("Option Button 1");
        r1.SetPosition(500, 100);
        r1.SetSize(100, 25);
        _ = ctrl.Group(r1);
        ExcelControlRadioButton? r2 = _ws.Drawings.AddRadioButtonControl("Option Button 2");
        r2.SetPosition(530, 100);
        r2.SetSize(100, 25);
        ExcelGroupShape g = ctrl.Group(r2);
        ExcelControlRadioButton? r3 = _ws.Drawings.AddRadioButtonControl("Option Button 3");
        r3.SetPosition(560, 100);
        r3.SetSize(100, 25);
        r3.FirstButton = true;
        g.Drawings.Add(r3);
    }

    [TestMethod]
    public void Group_ShapeAndChart()
    {
        _ws = _pck.Workbook.Worksheets.Add("ShapeAndChart");
        ExcelLineChart? chart = _ws.Drawings.AddLineChart("LineChart 1", eLineChartType.Line);

        ExcelShape? shape = _ws.Drawings.AddShape("Shape 1", eShapeStyle.Octagon);
        shape.SetPosition(200, 200);

        _ = chart.Group(shape);
    }

    [TestMethod]
    public void Group_PictureAndSlicer()
    {
        _ws = _pck.Workbook.Worksheets.Add("PictureAndSlicer");
        ExcelPicture? pic = _ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
        pic.SetPosition(400, 400);

        ExcelTable? tbl = _ws.Tables.Add(_ws.Cells["A1:B2"], "Table1");
        ExcelTableSlicer? slicer = tbl.Columns[0].AddSlicer();
        slicer.SetPosition(200, 200);

        _ = pic.Group(slicer);
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidOperationException))]
    public void Group_GroupIntoOthereWorksheetShouldFailText()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws1 = p.Workbook.Worksheets.Add("Sheet1");
        ExcelWorksheet? ws2 = p.Workbook.Worksheets.Add("Sheet2");
        ExcelControlGroupBox? ctrl1 = (ExcelControlGroupBox)ws1.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
        ctrl1.Text = "Groupbox 1";
        ctrl1.SetPosition(480, 80);
        ctrl1.SetSize(200, 120);

        ExcelControlGroupBox? ctrl2 = (ExcelControlGroupBox)ws1.Drawings.AddControl("GroupBox 2", eControlType.GroupBox);
        ctrl2.Text = "Groupbox 2";
        ctrl2.SetPosition(480, 400);
        ctrl2.SetSize(200, 120);

        ExcelControlRadioButton? r1 = ws2.Drawings.AddRadioButtonControl("Option Button 1");
        r1.SetPosition(500, 100);
        r1.SetSize(100, 25);
        _ = ctrl1.Group(r1);
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidOperationException))]
    public void Group_GroupIntoOtherGroupShouldFailTest()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("UnGroupAllDrawings");
        ExcelControlGroupBox? ctrl1 = (ExcelControlGroupBox)ws.Drawings.AddControl("GroupBox 1", eControlType.GroupBox);
        ctrl1.Text = "Groupbox 1";
        ctrl1.SetPosition(480, 80);
        ctrl1.SetSize(200, 120);

        ExcelControlGroupBox? ctrl2 = (ExcelControlGroupBox)ws.Drawings.AddControl("GroupBox 2", eControlType.GroupBox);
        ctrl2.Text = "Groupbox 2";
        ctrl2.SetPosition(480, 400);
        ctrl2.SetSize(200, 120);

        ExcelControlRadioButton? r1 = ws.Drawings.AddRadioButtonControl("Option Button 1");
        r1.SetPosition(500, 100);
        r1.SetSize(100, 25);
        _ = ctrl1.Group(r1);

        _ = ctrl2.Group(r1);
    }
}