using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class SumXTests
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
    public void Cleanup()
    {
        this._package.Dispose();
    }

    [TestMethod]
    public void SumX2My2_TwoRanges()
    {
        this._sheet.Cells["A1"].Value = 5;
        this._sheet.Cells["A2"].Value = 6;
        this._sheet.Cells["A3"].Value = 7;
        this._sheet.Cells["B1"].Value = 3;
        this._sheet.Cells["B2"].Value = 4;
        this._sheet.Cells["B3"].Value = 2;
        this._sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,B1:B3)";
        this._sheet.Calculate();
        Assert.AreEqual(81d, this._sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void SumX2My2_RangeAndArray()
    {
        this._sheet.Cells["A1"].Value = 5;
        this._sheet.Cells["A2"].Value = 6;
        this._sheet.Cells["A3"].Value = 7;
        this._sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,{3;4;2})";
        this._sheet.Calculate();
        Assert.AreEqual(81d, this._sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void SumX2My2_NonMatchingLengths()
    {
        this._sheet.Cells["A1"].Value = 5;
        this._sheet.Cells["A2"].Value = 6;
        this._sheet.Cells["A3"].Value = 7;
        this._sheet.Cells["B1"].Value = 3;
        this._sheet.Cells["B2"].Value = 4;
        this._sheet.Cells["B3"].Value = 2;
        this._sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,B1:B2)";
        this._sheet.Calculate();
        Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), this._sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void SumX2My2_NonNumeric1()
    {
        this._sheet.Cells["A1"].Value = 5;
        this._sheet.Cells["A2"].Value = "6";
        this._sheet.Cells["A3"].Value = 7;
        this._sheet.Cells["B1"].Value = 3;
        this._sheet.Cells["B2"].Value = 4;
        this._sheet.Cells["B3"].Value = 2;
        this._sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,B1:B3)";
        this._sheet.Calculate();
        Assert.AreEqual(61d, this._sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void SumX2My2_NonNumeric2()
    {
        this._sheet.Cells["A1"].Value = 5;
        this._sheet.Cells["A2"].Value = "6";
        this._sheet.Cells["A3"].Value = 7;
        this._sheet.Cells["B1"].Value = 3;
        this._sheet.Cells["B2"].Value = 4;
        this._sheet.Cells["B3"].Value = "2";
        this._sheet.Cells["C1"].Formula = "SUMX2MY2(A1:A3,B1:B3)";
        this._sheet.Calculate();
        Assert.AreEqual(16d, this._sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void SumXmY2_TwoRanges()
    {
        this._sheet.Cells["A1"].Value = 5;
        this._sheet.Cells["A2"].Value = 6;
        this._sheet.Cells["A3"].Value = 7;
        this._sheet.Cells["B1"].Value = 3;
        this._sheet.Cells["B2"].Value = 4;
        this._sheet.Cells["B3"].Value = 2;
        this._sheet.Cells["C1"].Formula = "SUMXMY2(A1:A3,B1:B3)";
        this._sheet.Calculate();
        Assert.AreEqual(33d, this._sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void SumX2pY2_TwoRanges()
    {
        this._sheet.Cells["A1"].Value = 5;
        this._sheet.Cells["A2"].Value = 6;
        this._sheet.Cells["A3"].Value = 7;
        this._sheet.Cells["B1"].Value = 3;
        this._sheet.Cells["B2"].Value = 4;
        this._sheet.Cells["B3"].Value = 2;
        this._sheet.Cells["C1"].Formula = "SUMX2PY2(A1:A3,B1:B3)";
        this._sheet.Calculate();
        Assert.AreEqual(139d, this._sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void SeriesSumShouldReturnCorrectResult()
    {
        this._sheet.Cells["A1"].Formula = "SERIESSUM( 5, 1, 1, {1,1,1,1,1} )";

        this._sheet.Cells["B1"].Value = 3;
        this._sheet.Cells["B2"].Value = 4;
        this._sheet.Cells["B3"].Value = 2;
        this._sheet.Cells["C1"].Formula = "SeriesSum(2,1,1;B1:B3)";

        this._sheet.Calculate();

        Assert.AreEqual(3905d, this._sheet.Cells["A1"].Value, "First assert failed");
        Assert.AreEqual(38d, this._sheet.Cells["C1"].Value, "Second assert failed");
    }
}