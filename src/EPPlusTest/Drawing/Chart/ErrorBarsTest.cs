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
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using System.Drawing;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ErrorBarsTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ErrorBars.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ErrorBars_StdDev()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_StDev");
            LoadTestdata(ws);

            ExcelLineChart? chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
            ExcelLineChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.StandardDeviation);
            serie.ErrorBars.Direction = eErrorBarDirection.Y;
            serie.ErrorBars.Value = 14;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.StandardDeviation, serie.ErrorBars.ValueType);
            Assert.AreEqual(eErrorBarDirection.Y, serie.ErrorBars.Direction);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);
        }
        [TestMethod]
        public void ErrorBars_StdErr()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_StErr");
            LoadTestdata(ws);

            ExcelLineChart? chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            ExcelLineChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.StandardError);
            serie.ErrorBars.Direction = eErrorBarDirection.X;
            serie.ErrorBars.NoEndCap = true;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.StandardError, serie.ErrorBars.ValueType);
            Assert.AreEqual(eErrorBarDirection.X, serie.ErrorBars.Direction);
            Assert.AreEqual(true, serie.ErrorBars.NoEndCap);
        }
        [TestMethod]
        public void ErrorBars_Percentage()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_Percentage");
            LoadTestdata(ws);

            ExcelLineChart? chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            ExcelLineChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage);
            serie.ErrorBars.Value = 5;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Plus, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Percentage, serie.ErrorBars.ValueType);
            Assert.AreEqual(eErrorBarDirection.Y, serie.ErrorBars.Direction);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);
        }
        [TestMethod]
        public void ErrorBars_Fixed()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_Fixed");
            LoadTestdata(ws);

            ExcelLineChart? chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            ExcelLineChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.FixedValue);
            serie.ErrorBars.Value = 5.2;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Plus, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.FixedValue, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);
        }
        [TestMethod]
        public void ErrorBars_Custom()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_Custom");
            LoadTestdata(ws);

            ExcelLineChart? chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            ExcelLineChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Custom);
            serie.ErrorBars.Plus.ValuesSource = "{1}";
            serie.ErrorBars.Minus.FormatCode = "General";
            serie.ErrorBars.Minus.ValuesSource = "=A2:A15";
            serie.ErrorBars.Minus.FormatCode = "0";
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Plus, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
            Assert.AreEqual("A2:A15", serie.ErrorBars.Minus.ValuesSource);
        }
        [TestMethod]
        public void ErrorBars_Scatter()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBarScatter");
            LoadTestdata(ws);

            ExcelScatterChart? chart = ws.Drawings.AddScatterChart("ScatterChart1", eScatterChartType.XYScatter);
            ExcelScatterChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
            serie.ErrorBars.Plus.ValuesSource = "{1}";
            serie.ErrorBars.Minus.FormatCode = "General";
            serie.ErrorBars.Minus.ValuesSource = "=ErrorBarScatter!$A$2:$A$15";
            serie.ErrorBars.Minus.FormatCode = "0";

            serie.ErrorBarsX.Plus.ValuesSource = "{2}";
            serie.ErrorBarsX.Minus.FormatCode = "General";
            serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBarScatter!$A$2:$A$16";
            serie.ErrorBarsX.Minus.FormatCode = "0";

            chart.SetPosition(1, 0, 5, 0);

            Assert.IsNotNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);
            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
            Assert.AreEqual("ErrorBarScatter!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
            Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

            Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
            Assert.AreEqual("ErrorBarScatter!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
        }
        [TestMethod]
        public void ErrorBars_ReadScatter()
        {
            using (ExcelPackage? p1 = new ExcelPackage())
            {
                ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("ErrorBarsScatter");
                LoadTestdata(ws);

                ExcelScatterChart? chart = ws.Drawings.AddScatterChart("ScatterChart1", eScatterChartType.XYScatter);
                ExcelScatterChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
                serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
                serie.ErrorBars.Plus.ValuesSource = "{1}";
                serie.ErrorBars.Minus.FormatCode = "General";
                serie.ErrorBars.Minus.ValuesSource = "=ErrorBarsScatter!$A$2:$A$15";
                serie.ErrorBars.Minus.FormatCode = "0";

                serie.ErrorBarsX.Plus.ValuesSource = "{2}";
                serie.ErrorBarsX.Minus.FormatCode = "General";
                serie.ErrorBarsX.Minus.ValuesSource = "ErrorBarsScatter!$A$2:$A$16";
                serie.ErrorBarsX.Minus.FormatCode = "0";

                chart.SetPosition(1, 0, 5, 0);
                p1.Save();
                using(ExcelPackage? p2=new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets["ErrorBarsScatter"];

                    chart = ws.Drawings[0].As.Chart.ScatterChart;
                    serie = chart.Series[0];

                    Assert.IsNotNull(serie.ErrorBars);
                    Assert.IsNotNull(serie.ErrorBarsX);
                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
                    Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

                    Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBarsScatter!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
                    Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

                    Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBarsScatter!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
                }
            }
        }
        [TestMethod]
        public void ErrorBars_Bubble()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_Bubble");
            LoadTestdata(ws);

            ExcelBubbleChart? chart = ws.Drawings.AddBubbleChart("BubbleChart1", eBubbleChartType.Bubble);
            ExcelBubbleChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
            serie.ErrorBars.Plus.ValuesSource = "{1}";
            serie.ErrorBars.Minus.FormatCode = "General";
            serie.ErrorBars.Minus.ValuesSource = "=ErrorBar_Bubble!$A$2:$A$15";
            serie.ErrorBars.Minus.FormatCode = "0";

            serie.ErrorBarsX.Plus.ValuesSource = "{2}";
            serie.ErrorBarsX.Minus.FormatCode = "General";
            serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBar_Bubble!$A$2:$A$16";
            serie.ErrorBarsX.Minus.FormatCode = "0";

            chart.SetPosition(1, 0, 5, 0);

            Assert.IsNotNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);
            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
            Assert.AreEqual("ErrorBar_Bubble!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
            Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

            Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
            Assert.AreEqual("ErrorBar_Bubble!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
        }
        [TestMethod]
        public void ErrorBars_ReadBubble()
        {
            using (ExcelPackage? p1 = new ExcelPackage())
            {
                ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("ErrorBars");
                LoadTestdata(ws);

                ExcelBubbleChart? chart = ws.Drawings.AddBubbleChart("BubbleChart1", eBubbleChartType.Bubble);
                ExcelBubbleChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
                serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
                serie.ErrorBars.Plus.ValuesSource = "{1}";
                serie.ErrorBars.Minus.FormatCode = "General";
                serie.ErrorBars.Minus.ValuesSource = "=ErrorBars!$A$2:$A$15";
                serie.ErrorBars.Minus.FormatCode = "0";

                serie.ErrorBarsX.Plus.ValuesSource = "{2}";
                serie.ErrorBarsX.Minus.FormatCode = "General";
                serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBars!$A$2:$A$16";
                serie.ErrorBarsX.Minus.FormatCode = "0";

                chart.SetPosition(1, 0, 5, 0);
                p1.Save();
                using (ExcelPackage? p2 = new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets["ErrorBars"];

                    chart = ws.Drawings[0].As.Chart.BubbleChart;
                    serie = chart.Series[0];

                    Assert.IsNotNull(serie.ErrorBars);
                    Assert.IsNotNull(serie.ErrorBarsX);
                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
                    Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

                    Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBars!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
                    Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

                    Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBars!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
                }
            }
        }
        [TestMethod]
        public void ErrorBars_Area()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_Area");
            LoadTestdata(ws);

            ExcelAreaChart? chart = ws.Drawings.AddAreaChart("AreaChart1", eAreaChartType.Area);
            ExcelAreaChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
            serie.ErrorBars.Plus.ValuesSource = "{1}";
            serie.ErrorBars.Minus.FormatCode = "General";
            serie.ErrorBars.Minus.ValuesSource = "=ErrorBar_Area!$A$2:$A$15";
            serie.ErrorBars.Minus.FormatCode = "0";

            serie.ErrorBarsX.Plus.ValuesSource = "{2}";
            serie.ErrorBarsX.Minus.FormatCode = "General";
            serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBar_Area!$A$2:$A$16";
            serie.ErrorBarsX.Minus.FormatCode = "0";

            chart.SetPosition(1, 0, 5, 0);

            Assert.IsNotNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);
            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
            Assert.AreEqual("ErrorBar_Area!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
            Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

            Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
            Assert.AreEqual("ErrorBar_Area!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
        }
        [TestMethod]
        public void ErrorBars_ReadArea()
        {
            using (ExcelPackage? p1 = new ExcelPackage())
            {
                ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("ErrorBars");
                LoadTestdata(ws);

                ExcelAreaChart? chart = ws.Drawings.AddAreaChart("ScatterChart1", eAreaChartType.Area);
                ExcelAreaChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
                serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
                serie.ErrorBars.Plus.ValuesSource = "{1}";
                serie.ErrorBars.Minus.FormatCode = "General";
                serie.ErrorBars.Minus.ValuesSource = "=ErrorBars!$A$2:$A$15";
                serie.ErrorBars.Minus.FormatCode = "0";

                serie.ErrorBarsX.Plus.ValuesSource = "{2}";
                serie.ErrorBarsX.Minus.FormatCode = "General";
                serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBars!$A$2:$A$16";
                serie.ErrorBarsX.Minus.FormatCode = "0";

                chart.SetPosition(1, 0, 5, 0);
                p1.Save();
                using (ExcelPackage? p2 = new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets["ErrorBars"];

                    chart = ws.Drawings[0].As.Chart.AreaChart;
                    serie = chart.Series[0];

                    Assert.IsNotNull(serie.ErrorBars);
                    Assert.IsNotNull(serie.ErrorBarsX);
                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
                    Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

                    Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBars!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
                    Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

                    Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBars!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
                }
            }
        }
        [TestMethod]
        public void ErrorBars_Delete()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_Percentage_removed");
            LoadTestdata(ws);

            ExcelLineChart? chart = ws.Drawings.AddLineChart("LineChart1_DeletedErrorbars", eLineChartType.Line);
            ExcelLineChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage);
            serie.ErrorBars.Value = 5;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Plus, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Percentage, serie.ErrorBars.ValueType);
            Assert.AreEqual(eErrorBarDirection.Y, serie.ErrorBars.Direction);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            serie.ErrorBars.Remove();
            Assert.IsNull(serie.ErrorBars);
        }
        [TestMethod]
        public void ErrorBarsScatter_Delete()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ErrorBar_Scatter_removed");
            LoadTestdata(ws);

            ExcelScatterChart? chart = ws.Drawings.AddScatterChart("LineChart1_DeletedErrorbars", eScatterChartType.XYScatter);
            ExcelScatterChartSerie? serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage);
            Assert.IsNotNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);

            serie.ErrorBars.Remove();
            Assert.IsNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);

            serie.ErrorBarsX.Remove();
            Assert.IsNull(serie.ErrorBars);
            Assert.IsNull(serie.ErrorBarsX);

        }

    }
}
