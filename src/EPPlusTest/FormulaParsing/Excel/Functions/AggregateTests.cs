using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions;

[TestClass]
public class AggregateTests
{
    private ExcelPackage _package;
    private ExcelWorksheet _sheet;

    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        this._sheet = this._package.Workbook.Worksheets.Add("test");
    }

    [TestCleanup]
    public void Cleanup() => this._package.Dispose();

    private void LoadData1()
    {
        this._sheet.Cells["A1"].Value = 3;
        this._sheet.Cells["A2"].Value = 2.5;
        this._sheet.Cells["A3"].Value = 1;
        this._sheet.Cells["A4"].Value = 6;
        this._sheet.Cells["A5"].Value = -2;
    }

    // Tests for Ignore nothing

    [TestMethod]
    public void ShouldHandleAverage()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 1, 4, A1, A2, A3, A4, A5 )";
        this._sheet.Calculate();
        Assert.AreEqual(2.1d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void ShouldHandleSum()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 9, 4, A1, A2, A3, A4, A5 )";
        this._sheet.Calculate();
        Assert.AreEqual(10.5d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void ShouldHandleMin()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 5, 4, A1:A5 )";
        this._sheet.Calculate();
        Assert.AreEqual(-2d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void ShouldHandleLarge()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 14, 4, A1:A5, 2 )";
        this._sheet.Calculate();
        Assert.AreEqual(3d, this._sheet.Cells["A6"].Value);
    }

    // Tests for Ignore hidden cells

    [TestMethod]
    public void HiddenCells_ShouldHandleAverage()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 1, 5, A1, A2, A3, A4, A5 )";
        this._sheet.Row(3).Hidden = true;
        this._sheet.Calculate();
        Assert.AreEqual(2.375d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void HiddenCells_ShouldHandleSum()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 9, 5, A1, A2, A3, A4, A5 )";
        this._sheet.Row(3).Hidden = true;
        this._sheet.Calculate();
        Assert.AreEqual(9.5d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void HiddenCells_ShouldHandleMin()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 5, 5, A1, A2, A3, A4, A5 )";
        this._sheet.Row(5).Hidden = true;
        this._sheet.Calculate();
        Assert.AreEqual(1d, this._sheet.Cells["A6"].Value);
    }

    // Tests for ignoring errors

    [TestMethod]
    public void Errors_ShouldHandleAverage()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 1, 6, A1, A2, A3, A4, A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(2.375d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A3"].Value = 1;
        this._sheet.Calculate();
        Assert.AreEqual(2.1d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleCount()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 2, 6, A1, A2, A3, A4, A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(4d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A3"].Value = 1;
        this._sheet.Calculate();
        Assert.AreEqual(5d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleCountA()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 3, 6, A1, A2, A3, A4, A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(4d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A3"].Value = 1;
        this._sheet.Calculate();
        Assert.AreEqual(5d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleMax()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 4, 6, A1, A2, A3, A4, A5 )";
        this._sheet.Cells["A4"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(3d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A6"].Formula = "AGGREGATE( 4, 4, A1:A5 )";
        this._sheet.Calculate();
        Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleMin()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 5, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(-2d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A6"].Formula = "AGGREGATE( 5, 4, A1:A5 )";
        this._sheet.Calculate();
        Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleProduct()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 6, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(-90d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A6"].Formula = "AGGREGATE( 5, 4, A1:A5 )";
        this._sheet.Calculate();
        Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleStdevS()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 7, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        double result = (double)this._sheet.Cells["A6"].Value;
        result = System.Math.Round(result, 5);
        Assert.AreEqual(3.30088d, result);

        this._sheet.Cells["A6"].Formula = "AGGREGATE( 7, 4, A1:A5 )";
        this._sheet.Calculate();
        Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleStdevP()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 8, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        double result = (double)this._sheet.Cells["A6"].Value;
        result = System.Math.Round(result, 5);
        Assert.AreEqual(2.85865d, result);

        this._sheet.Cells["A6"].Formula = "AGGREGATE( 8, 4, A1:A5 )";
        this._sheet.Calculate();
        Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleSum()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 9, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(9.5d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A3"].Value = 1;
        this._sheet.Calculate();
        Assert.AreEqual(10.5d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleVarS()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 10, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        double result = (double)this._sheet.Cells["A6"].Value;
        result = System.Math.Round(result, 5);
        Assert.AreEqual(10.89583d, result);

        this._sheet.Cells["A3"].Value = 1;
        this._sheet.Calculate();
        Assert.AreEqual(8.55d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleVarP()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 11, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(8.171875d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A3"].Value = 1;
        this._sheet.Calculate();
        double result = (double)this._sheet.Cells["A6"].Value;
        result = System.Math.Round(result, 2);
        Assert.AreEqual(6.84d, result);
    }

    [TestMethod]
    public void Errors_ShouldHandleMedian()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 12, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(2.75d, this._sheet.Cells["A6"].Value);

        this._sheet.Cells["A3"].Value = 1;
        this._sheet.Calculate();
        double result = (double)this._sheet.Cells["A6"].Value;
        result = System.Math.Round(result, 2);
        Assert.AreEqual(2.5d, result);
    }

    [TestMethod]
    public void Errors_ShouldHandleModeSngl()
    {
        this.LoadData1();
        this._sheet.Cells["A2"].Value = 3;
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 13, 6, A1:A5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(3d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleLarge()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 14, 6, A1:A5, 1 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(6d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleSmall()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 15, 6, A1:A5, 1 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(-2d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandlePercentileInc()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 16, 6, A1:A5, 0 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(-2d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleQuartileInc()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 17, 6, A1:A5, 1 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(1.375d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandlePercentileExc()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 18, 6, A1:A5, 0.5 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(2.75d, this._sheet.Cells["A6"].Value);
    }

    [TestMethod]
    public void Errors_ShouldHandleQuartileExc()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 19, 6, A1:A5, 1 )";
        this._sheet.Cells["A3"].Formula = "1/0";
        this._sheet.Calculate();
        Assert.AreEqual(-0.875d, this._sheet.Cells["A6"].Value);
    }

    // Tests for ignoring nested aggregate functions

    [TestMethod]
    public void IngoreNestedAggregateFunction()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 19, 6, A1:A5, 1)";
        this._sheet.Cells["A7"].Formula = "AGGREGATE( 2, 0, A1:A6)";
        this._sheet.Calculate();
        Assert.AreEqual(5d, this._sheet.Cells["A7"].Value);
    }

    [TestMethod]
    public void IncludeNestedAggregateFunction()
    {
        this.LoadData1();
        this._sheet.Cells["A6"].Formula = "AGGREGATE( 19, 6, A1:A5, 1)";
        this._sheet.Cells["A7"].Formula = "AGGREGATE( 2, 4, A1:A6)";
        this._sheet.Calculate();
        Assert.AreEqual(6d, this._sheet.Cells["A7"].Value);
    }

    [TestMethod]
    public void ShouldHandleMultipleLevelsOfAggregate()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet3 = package.Workbook.Worksheets.Add("sheet3");
        sheet3.Cells["A1"].Value = 26959.64;
        sheet3.Cells["A2"].Value = 82272d;
        sheet3.Cells["A3"].Formula = "AGGREGATE(9,0,A1:A2)";
        sheet3.Cells["A4"].Formula = "AGGREGATE(9,0,A1:A3)";

        ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("sheet2");
        sheet2.Cells["A1"].Formula = "sheet3!A4";
        package.Workbook.Calculate();
        Assert.AreEqual(109231.64d, sheet2.Cells["A1"].Value);

        sheet3.Cells["A3"].Formula = "AGGREGATE(8,0,A1:A2)";
        sheet3.Cells["A4"].Formula = "AGGREGATE(8,0,A1:A3)";
        package.Workbook.Calculate();
        Assert.AreEqual(27656.18, sheet2.Cells["A1"].Value);

        sheet3.Cells["A3"].Formula = "AGGREGATE(7,0,A1:A2)";
        sheet3.Cells["A4"].Formula = "AGGREGATE(7,0,A1:A3)";
        package.Workbook.Calculate();
        Assert.AreEqual(39111.7448d, System.Math.Round((double)sheet2.Cells["A1"].Value, 4));
    }
}