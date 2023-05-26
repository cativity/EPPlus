using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Chart;

[TestClass]
public class ChartSeriesTest : TestBase
{
    static ExcelPackage _pck;
    [ClassInitialize]
    public static void Init(TestContext context)
    {
        _pck = OpenPackage("ChartSingleSerie.xlsx", true);
    }
    [ClassCleanup]
    public static void Cleanup()
    {
        SaveAndCleanup(_pck);
    }
    #region Single Serie
    [TestMethod]
    public void AddSunburstChartSingleSerie()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Sunburst");
        LoadHierarkiTestData(ws);
        ExcelSunburstChart? chart = ws.Drawings.AddSunburstChart("Sunburst1");
        ExcelChartExSerie? serie = chart.Series.Add(ws.Cells["D2:D17"]);
        chart.SetPosition(2, 0, 15, 0);
        chart.SetSize(1600, 900);
        serie.DataLabel.Position = eLabelPosition.Center;
        serie.DataLabel.ShowCategory = true;
        serie.DataLabel.ShowValue = true;
        ExcelChartExDataPoint? dp = serie.DataPoints.Add(2);

        Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
        Assert.IsInstanceOfType(chart, typeof(ExcelSunburstChart));
        Assert.AreEqual(0, chart.Axis.Length);
        Assert.IsNull(chart.XAxis);
        Assert.IsNull(chart.YAxis);

    }
    [TestMethod]
    public void AddPieChartSingleSerie()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Pie");
        LoadHierarkiTestData(ws);
        ExcelPieChart? chart = ws.Drawings.AddPieChart("Pie1", ePieChartType.Pie);
        ExcelPieChartSerie? serie = chart.Series.Add(ws.Cells["D2:D17"]);
        chart.SetPosition(2, 0, 15, 0);
        chart.SetSize(1600, 900);
        serie.DataLabel.Position = eLabelPosition.Center;
        serie.DataLabel.ShowCategory = true;
        serie.DataLabel.ShowValue = true;
        ExcelChartDataPoint? dp = serie.DataPoints.Add(2);

        Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
        Assert.IsInstanceOfType(chart, typeof(ExcelPieChart));
        Assert.AreEqual(0, chart.Axis.Length);
        Assert.IsNull(chart.XAxis);
        Assert.IsNull(chart.YAxis);

    }
    [TestMethod]
    public void AddColumnChartSingleSerieWithSecondSerieWithCategory()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Column");
        LoadHierarkiTestData(ws);
        ExcelBarChart? chart = ws.Drawings.AddBarChart("Bar1", eBarChartType.Column3D);
        ExcelBarChartSerie? serie1 = chart.Series.Add(ws.Cells["D2:D17"]);
        ExcelBarChartSerie? serie2 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["C2:C17"]);
        chart.SetPosition(2, 0, 15, 0);
        chart.SetSize(1600, 900);

        Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie1.Series);
        Assert.AreEqual("", serie1.XSeries);
        Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie2.Series);
        Assert.AreEqual(ws.Cells["C2:C17"].FullAddressAbsolute, serie2.XSeries);

        Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
        Assert.IsInstanceOfType(chart, typeof(ExcelBarChart));
        Assert.AreEqual(2, chart.Axis.Length);
        Assert.IsNotNull(chart.XAxis);
        Assert.IsNotNull(chart.YAxis);

    }

    [TestMethod]
    public void AddColumnChartSingleSerieWithSecondSerieWithCategoryWithLinear()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ColumnWithinLinear");
        LoadHierarkiTestData(ws);
        ExcelBarChart? chart = ws.Drawings.AddBarChart("Bar1", eBarChartType.Column3D);

        //Change chart colorMethod from Cylce to WithinLinear
        chart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyleMultiSeries.Column3dChartStyle1,
                                         OfficeOpenXml.Drawing.Chart.Style.ePresetChartColors.ColorfulPalette1);
        chart.StyleManager.ColorsManager.Method = OfficeOpenXml.Drawing.Chart.Style.eChartColorStyleMethod.WithinLinear;

        //make series only have range of 1 so that the serie2(index=1) is the same as the number of cells in the range
        //which causes System.ArgumentException: Negative percentage not allowed
        ExcelBarChartSerie? serie1 = chart.Series.Add(ws.Cells["D2"]);
        ExcelBarChartSerie? serie2 = chart.Series.Add(ws.Cells["D2"], ws.Cells["C2"]);
    }

    [TestMethod]
    public void AddChartWithLegendEntries()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("LegendEntries");
        LoadHierarkiTestData(ws);
        ExcelBarChart? chart = ws.Drawings.AddBarChart("Bar1", eBarChartType.Column3D);
        ExcelBarChartSerie? serie1 = chart.Series.Add(ws.Cells["D2:D17"]);
        ExcelBarChartSerie? serie2 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["C2:C17"]);
        ExcelBarChartSerie? serie3 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["B2:B17"]);

        serie1.Header = "Serie 1";
        serie2.Header = "Serie 2-Deleted";
        serie3.Header = "Serie 3-Font Changed";

        chart.Legend.Entries[2].Font.Fill.Style = eFillStyle.SolidFill;
        chart.Legend.Entries[2].Font.Fill.SolidFill.Color.SetRgbColor(Color.Red);

        chart.Legend.Entries[1].Deleted = true;

        chart.SetPosition(2, 0, 15, 0);
        chart.SetSize(1600, 900);

        Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie1.Series);
        Assert.AreEqual("", serie1.XSeries);
        Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie2.Series);
        Assert.AreEqual(ws.Cells["C2:C17"].FullAddressAbsolute, serie2.XSeries);

        Assert.AreEqual("Serie 1", serie1.Header);
        Assert.AreEqual("Serie 2-Deleted", serie2.Header);
        Assert.IsTrue(chart.Legend.Entries[1].Deleted);
        Assert.AreEqual("Serie 3-Font Changed", serie3.Header);

        Assert.AreEqual(eFillStyle.SolidFill,chart.Legend.Entries[2].Font.Fill.Style);
        Assert.AreEqual(Color.Red.ToArgb(), chart.Legend.Entries[2].Font.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
    }

    [TestMethod]
    public void ReadChartWithLegendEntries()
    {
        using ExcelPackage? p1 = new ExcelPackage();
        ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("LegendEntries");
        LoadHierarkiTestData(ws);
        ExcelBarChart? chart = ws.Drawings.AddBarChart("Bar1", eBarChartType.Column3D);
        ExcelBarChartSerie? serie1 = chart.Series.Add(ws.Cells["D2:D17"]);
        ExcelBarChartSerie? serie2 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["C2:C17"]);
        ExcelBarChartSerie? serie3 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["B2:B17"]);

        serie1.Header = "Serie 1";
        serie2.Header = "Serie 2-Deleted";
        serie3.Header = "Serie 3-Font Changed";

        chart.Legend.Entries[2].Font.Fill.Style = eFillStyle.SolidFill;
        chart.Legend.Entries[2].Font.Fill.SolidFill.Color.SetRgbColor(Color.Red);

        chart.Legend.Entries[1].Deleted = true;

        chart.SetPosition(2, 0, 15, 0);
        chart.SetSize(1600, 900);

        //Assert p1
        Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie1.Series);
        Assert.AreEqual("", serie1.XSeries);
        Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie2.Series);
        Assert.AreEqual(ws.Cells["C2:C17"].FullAddressAbsolute, serie2.XSeries);

        Assert.AreEqual("Serie 1", serie1.Header);
        Assert.AreEqual("Serie 2-Deleted", serie2.Header);
        Assert.IsTrue(chart.Legend.Entries[1].Deleted);
        Assert.AreEqual("Serie 3-Font Changed", serie3.Header);

        Assert.AreEqual(eFillStyle.SolidFill, chart.Legend.Entries[2].Font.Fill.Style);
        Assert.AreEqual(Color.Red.ToArgb(), chart.Legend.Entries[2].Font.Fill.SolidFill.Color.RgbColor.Color.ToArgb());

        p1.Save();

        using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
        ws = p2.Workbook.Worksheets[0];
        chart = ws.Drawings[0].As.Chart.BarChart;
        serie1 = chart.Series[0];
        serie2 = chart.Series[1];
        serie3 = chart.Series[2];

        //Assert p2 
        Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie1.Series);
        Assert.AreEqual("", serie1.XSeries);
        Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie2.Series);
        Assert.AreEqual(ws.Cells["C2:C17"].FullAddressAbsolute, serie2.XSeries);

        Assert.AreEqual("Serie 1", serie1.Header);
        Assert.AreEqual("Serie 2-Deleted", serie2.Header);
        Assert.IsTrue(chart.Legend.Entries[1].Deleted);
        Assert.AreEqual("Serie 3-Font Changed", serie3.Header);

        Assert.IsFalse(chart.Legend.Entries[0].Deleted);
        Assert.AreEqual(eFillStyle.SolidFill, chart.Legend.Entries[2].Font.Fill.Style);
        Assert.AreEqual(Color.Red.ToArgb(), chart.Legend.Entries[2].Font.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
    }
    #endregion
}