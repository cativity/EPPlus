﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace EPPlusTest.Drawing.Chart;

[TestClass]
public class ChartTitleTests : TestBase
{
    static ExcelPackage _pck;
    static ExcelWorksheet _wsData;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        _pck = OpenPackage("ChartTitle.xlsx", true);
        _wsData = _pck.Workbook.Worksheets.Add("Data");
        LoadItemData(_wsData);
    }

    [ClassCleanup]
    public static void Cleanup()
    {
        string? dirName = _pck.File.DirectoryName;
        string? fileName = _pck.File.FullName;
        SaveAndCleanup(_pck);

        if (File.Exists(fileName))
        {
            File.Copy(fileName, dirName + "\\ChartTitleRead.xlsx", true);
        }
    }

    [TestMethod]
    public void AddLineChartWithTextTitle()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LineChartTextTitle");
        ExcelLineChart? chart = ws.Drawings.AddLineChart("lineChart1", eLineChartType.Line);
        chart.Title.Text = "LineChart - Text";
        chart.Series.Add("Data!N1:N10", "Data!K1:K10");
    }

    [TestMethod]
    public void AddLineChartWithCellLinkTitle()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LineChartCellLinkTitle");
        ExcelLineChart? chart = ws.Drawings.AddLineChart("lineChart2", eLineChartType.Line);
        ws.Cells["A1"].Value = "Linked Cell Title";
        chart.Title.LinkedCell = ws.Cells["A1"];
        chart.Series.Add("Data!N1:N10", "Data!K1:K10");
        Assert.AreEqual("Linked Cell Title", chart.Title.Text);
    }

    [TestMethod]
    public void AddLineChart_With_Text_Then_CellLink_Title()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LineChartLinkTitleOverride");
        ExcelLineChart? chart = ws.Drawings.AddLineChart("lineChart3", eLineChartType.Line);
        chart.Title.Text = "Line Chart - Text";
        _wsData.Cells["A1"].Value = "Linked Cell Title-DataSheet";
        chart.Title.LinkedCell = _wsData.Cells["A1"];
        chart.Series.Add("Data!N1:N10", "Data!K1:K10");
        Assert.AreEqual("Linked Cell Title-DataSheet", chart.Title.Text);
    }

    [TestMethod]
    public void AddLineChart_With_CellLink_Then_Text_Title()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LineTextTitleOverride");
        ExcelLineChart? chart = ws.Drawings.AddLineChart("lineChart4", eLineChartType.Line);
        ws.Cells["A1"].Value = "Linked Cell Title";
        chart.Title.LinkedCell = ws.Cells["A1"];
        chart.Title.Text = "Line Chart - Text";
        chart.Series.Add("Data!N1:N10", "Data!K1:K10");
        Assert.AreEqual("Line Chart - Text", chart.Title.Text);
    }

    [TestMethod]
    public void AddBarChartWithAxisTextTitle()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("BarChartTextTitle");
        ExcelBarChart? chart = ws.Drawings.AddBarChart("barChart1", eBarChartType.BarClustered);
        chart.Series.Add("Data!N1:N10", "Data!K1:K10");
        chart.XAxis.AddTitle("Linked Cell Title");
    }

    [TestMethod]
    public void AddBarChartWithAxisLinkedTitle()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("BarChartLinkedCellTitle");
        ExcelBarChart? chart = ws.Drawings.AddBarChart("barChart1", eBarChartType.BarClustered);
        ws.Cells["A1"].Value = "Linked Cell Title";
        chart.Series.Add("Data!N1:N10", "Data!K1:K10");
        chart.XAxis.AddTitle(ws.Cells["A1"]);
    }
}