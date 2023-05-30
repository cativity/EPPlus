using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering.Convert;

[TestClass]
public class ConvertTests
{
    private ExcelPackage _package;
    //private EpplusExcelDataProvider _provider;
    private ParsingContext _parsingContext;
    private ExcelWorksheet _worksheet;

    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        //this._provider = new EpplusExcelDataProvider(this._package);
        this._parsingContext = ParsingContext.Create();
        _ = this._parsingContext.Scopes.NewScope(RangeAddress.Empty);
        this._worksheet = this._package.Workbook.Worksheets.Add("testsheet");
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._package.Dispose();
    }

    [TestMethod]
    public void Bin2Dec_Tests()
    {
        this._worksheet.Cells["A1"].Value = "10";
        this._worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(2, this._worksheet.Cells["A2"].Value, "10 was not 2");

        this._worksheet.Cells["A1"].Value = "11";
        this._worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(3, this._worksheet.Cells["A2"].Value, "11 was not 3");

        this._worksheet.Cells["A1"].Value = "0000000010";
        this._worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(2, this._worksheet.Cells["A2"].Value, "0000000010 was not 2");

        this._worksheet.Cells["A1"].Value = "1111111110";
        this._worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(-2, this._worksheet.Cells["A2"].Value, "1111111110 was not -2");

        this._worksheet.Cells["A1"].Value = 110;
        this._worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(6, this._worksheet.Cells["A2"].Value, "110 was not 6");

        this._worksheet.Cells["A1"].Value = 1110000110;
        this._worksheet.Cells["A2"].Formula = "BIN2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(-122, this._worksheet.Cells["A2"].Value, "110 was not 6");
    }

    [TestMethod]
    public void Bin2Hex_Tests()
    {
        this._worksheet.Cells["A1"].Value = "10";
        this._worksheet.Cells["A2"].Formula = "BIN2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("2", this._worksheet.Cells["A2"].Value, "10 was not 2");

        this._worksheet.Cells["A1"].Value = "0000000001";
        this._worksheet.Cells["A2"].Formula = "BIN2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1", this._worksheet.Cells["A2"].Value, "0000000001 was not 1 but " + this._worksheet.Cells["A2"].Value);

        this._worksheet.Cells["A1"].Value = "1111111110";
        this._worksheet.Cells["A2"].Formula = "BIN2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("FFFFFFFFFE", this._worksheet.Cells["A2"].Value, "1111111110 was not FFFFFFFFFE but " + this._worksheet.Cells["A2"].Value);

        this._worksheet.Cells["A1"].Value = "11101";
        this._worksheet.Cells["A2"].Formula = "BIN2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1D", this._worksheet.Cells["A2"].Value, "11101 was not 1D but " + this._worksheet.Cells["A2"].Value);

        this._worksheet.Cells["A1"].Value = "10";
        this._worksheet.Cells["A2"].Formula = "BIN2HEX(A1,10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000002", this._worksheet.Cells["A2"].Value, "10 (padded with 10) was not 0000000002 but " + this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void Bin2Oct_Tests()
    {
        this._worksheet.Cells["A1"].Value = "101";
        this._worksheet.Cells["A2"].Formula = "BIN2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("5", this._worksheet.Cells["A2"].Value, "101 was not 5");

        this._worksheet.Cells["A1"].Value = "0000000001";
        this._worksheet.Cells["A2"].Formula = "BIN2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1", this._worksheet.Cells["A2"].Value, "0000000001 was not 1");

        this._worksheet.Cells["A1"].Value = "10";
        this._worksheet.Cells["A2"].Formula = "BIN2OCT(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000002", this._worksheet.Cells["A2"].Value, "10 was not 0000000002");

        this._worksheet.Cells["A1"].Value = "1111111110";
        this._worksheet.Cells["A2"].Formula = "BIN2OCT(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("7777777776", this._worksheet.Cells["A2"].Value, "1111111110 was not 7777777776");

        this._worksheet.Cells["A1"].Value = "1110";
        this._worksheet.Cells["A2"].Formula = "BIN2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("16", this._worksheet.Cells["A2"].Value, "1110 was not 16");
    }

    [TestMethod]
    public void Dec2Bin_Tests()
    {
        this._worksheet.Cells["A1"].Value = "2";
        this._worksheet.Cells["A2"].Formula = "DEC2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("10", this._worksheet.Cells["A2"].Value, "2 was not 10");

        this._worksheet.Cells["A1"].Value = "3";
        this._worksheet.Cells["A2"].Formula = "DEC2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("11", this._worksheet.Cells["A2"].Value, "3 was not 11");

        this._worksheet.Cells["A1"].Value = "2";
        this._worksheet.Cells["A2"].Formula = "DEC2BIN(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000010", this._worksheet.Cells["A2"].Value, "2 (padded with 10) was not 0000000010");

        this._worksheet.Cells["A1"].Value = "-2";
        this._worksheet.Cells["A2"].Formula = "DEC2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1111111110", this._worksheet.Cells["A2"].Value, "-2 was not 1111111110");

        this._worksheet.Cells["A1"].Value = "6";
        this._worksheet.Cells["A2"].Formula = "DEC2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("110", this._worksheet.Cells["A2"].Value, "6 was not 110");
    }

    [TestMethod]
    public void Dec2Hex_Tests()
    {
        this._worksheet.Cells["A1"].Value = "10";
        this._worksheet.Cells["A2"].Formula = "DEC2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("A", this._worksheet.Cells["A2"].Value, "10 was not A");

        this._worksheet.Cells["A1"].Value = "31";
        this._worksheet.Cells["A2"].Formula = "DEC2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1F", this._worksheet.Cells["A2"].Value, "31 was not 1F");

        this._worksheet.Cells["A1"].Value = "16";
        this._worksheet.Cells["A2"].Formula = "DEC2HEX(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000010", this._worksheet.Cells["A2"].Value, "16 was not 0000000010");

        this._worksheet.Cells["A1"].Value = "-16";
        this._worksheet.Cells["A2"].Formula = "DEC2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("FFFFFFFFF0", this._worksheet.Cells["A2"].Value, "-16 was not FFFFFFFFF0");

        this._worksheet.Cells["A1"].Value = "273";
        this._worksheet.Cells["A2"].Formula = "DEC2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("111", this._worksheet.Cells["A2"].Value, "273 was not 111");
    }

    [TestMethod]
    public void Dec2Oct_Tests()
    {
        this._worksheet.Cells["A1"].Value = 8;
        this._worksheet.Cells["A2"].Formula = "DEC2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("10", this._worksheet.Cells["A2"].Value, "8 was not 10");

        this._worksheet.Cells["A1"].Value = 18;
        this._worksheet.Cells["A2"].Formula = "DEC2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("22", this._worksheet.Cells["A2"].Value, "18 was not 22");

        this._worksheet.Cells["A1"].Value = 8;
        this._worksheet.Cells["A2"].Formula = "DEC2OCT(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000010", this._worksheet.Cells["A2"].Value, "8 was not 0000000010");

        this._worksheet.Cells["A1"].Value = -8;
        this._worksheet.Cells["A2"].Formula = "DEC2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("7777777770", this._worksheet.Cells["A2"].Value, "-8 was not 7777777770");

        this._worksheet.Cells["A1"].Value = 237;
        this._worksheet.Cells["A2"].Formula = "DEC2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("355", this._worksheet.Cells["A2"].Value, "237 was not 355");
    }

    [TestMethod]
    public void Hex2Bin_Tests()
    {
        this._worksheet.Cells["A1"].Value = "2";
        this._worksheet.Cells["A2"].Formula = "HEX2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("10", this._worksheet.Cells["A2"].Value, "2 was not 10");

        this._worksheet.Cells["A1"].Value = "0000000001";
        this._worksheet.Cells["A2"].Formula = "HEX2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1", this._worksheet.Cells["A2"].Value, "0000000001 was not 1");

        this._worksheet.Cells["A1"].Value = "2";
        this._worksheet.Cells["A2"].Formula = "HEX2BIN(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000010", this._worksheet.Cells["A2"].Value, "2 was not 0000000010");

        this._worksheet.Cells["A1"].Value = "FFFFFFFF9C";
        this._worksheet.Cells["A2"].Formula = "HEX2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1110011100", this._worksheet.Cells["A2"].Value, "FFFFFFFF9C was not 1110011100");

        this._worksheet.Cells["A1"].Value = "F0";
        this._worksheet.Cells["A2"].Formula = "HEX2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("11110000", this._worksheet.Cells["A2"].Value, "F0 was not 11110000");

        this._worksheet.Cells["A1"].Value = "1D";
        this._worksheet.Cells["A2"].Formula = "HEX2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("11101", this._worksheet.Cells["A2"].Value, "1D was not 11101");
    }

    [TestMethod]
    public void Hex2Dec_Tests()
    {
        this._worksheet.Cells["A1"].Value = "A";
        this._worksheet.Cells["A2"].Formula = "HEX2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(10d, this._worksheet.Cells["A2"].Value, "A was not 10");

        this._worksheet.Cells["A1"].Value = "1F";
        this._worksheet.Cells["A2"].Formula = "HEX2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(31d, this._worksheet.Cells["A2"].Value, "1F was not 31");

        this._worksheet.Cells["A1"].Value = "0000000010";
        this._worksheet.Cells["A2"].Formula = "HEX2DEC(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(16d, this._worksheet.Cells["A2"].Value, "0000000010 was not 16");

        this._worksheet.Cells["A1"].Value = "FFFFFFFFF0";
        this._worksheet.Cells["A2"].Formula = "HEX2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(-16d, this._worksheet.Cells["A2"].Value, "FFFFFFFFF0 was not -16");

        this._worksheet.Cells["A1"].Value = "FFFFFFFF10";
        this._worksheet.Cells["A2"].Formula = "HEX2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(-240d, this._worksheet.Cells["A2"].Value, "FFFFFFFF10 was not -240");

        this._worksheet.Cells["A1"].Value = "111";
        this._worksheet.Cells["A2"].Formula = "HEX2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(273d, this._worksheet.Cells["A2"].Value, "111 was not 273");
    }

    [TestMethod]
    public void Hex2Oct_Tests()
    {
        this._worksheet.Cells["A1"].Value = "A";
        this._worksheet.Cells["A2"].Formula = "HEX2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("12", this._worksheet.Cells["A2"].Value, "A was not 12");

        this._worksheet.Cells["A1"].Value = "000000000F";
        this._worksheet.Cells["A2"].Formula = "HEX2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("17", this._worksheet.Cells["A2"].Value, "000000000F was not 17");

        this._worksheet.Cells["A1"].Value = "8";
        this._worksheet.Cells["A2"].Formula = "HEX2OCT(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000010", this._worksheet.Cells["A2"].Value, "8 was not 0000000010");

        this._worksheet.Cells["A1"].Value = "FFFFFFFFF8";
        this._worksheet.Cells["A2"].Formula = "HEX2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("7777777770", this._worksheet.Cells["A2"].Value, "FFFFFFFFF0 was not 7777777770");

        this._worksheet.Cells["A1"].Value = "1F3";
        this._worksheet.Cells["A2"].Formula = "HEX2OCT(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("763", this._worksheet.Cells["A2"].Value, "1F3 was not 273");
    }

    [TestMethod]
    public void Oct2Bin_Tests()
    {
        this._worksheet.Cells["A1"].Value = "5";
        this._worksheet.Cells["A2"].Formula = "OCT2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("101", this._worksheet.Cells["A2"].Value, "5 was not 101");

        this._worksheet.Cells["A1"].Value = "0000000001";
        this._worksheet.Cells["A2"].Formula = "OCT2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1", this._worksheet.Cells["A2"].Value, "0000000001 was not 1");

        this._worksheet.Cells["A1"].Value = "2";
        this._worksheet.Cells["A2"].Formula = "OCT2BIN(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000010", this._worksheet.Cells["A2"].Value, "2 was not 0000000010");

        this._worksheet.Cells["A1"].Value = "7777777770";
        this._worksheet.Cells["A2"].Formula = "OCT2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1111111000", this._worksheet.Cells["A2"].Value, "7777777770 was not 1111111000");

        this._worksheet.Cells["A1"].Value = "16";
        this._worksheet.Cells["A2"].Formula = "OCT2BIN(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1110", this._worksheet.Cells["A2"].Value, "1F3 was not 1110");
    }

    [TestMethod]
    public void Oct2Dec_Tests()
    {
        this._worksheet.Cells["A1"].Value = "10";
        this._worksheet.Cells["A2"].Formula = "OCT2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(8d, this._worksheet.Cells["A2"].Value, "10 was not 8");

        this._worksheet.Cells["A1"].Value = "22";
        this._worksheet.Cells["A2"].Formula = "OCT2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(18d, this._worksheet.Cells["A2"].Value, "22 was not 18");

        this._worksheet.Cells["A1"].Value = "0000000010";
        this._worksheet.Cells["A2"].Formula = "OCT2DEC(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(8d, this._worksheet.Cells["A2"].Value, "0000000010 was not 8");

        this._worksheet.Cells["A1"].Value = "7777777770";
        this._worksheet.Cells["A2"].Formula = "OCT2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(-8d, this._worksheet.Cells["A2"].Value, "7777777770 was not -8");

        this._worksheet.Cells["A1"].Value = "355";
        this._worksheet.Cells["A2"].Formula = "OCT2DEC(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(237d, this._worksheet.Cells["A2"].Value, "355 was not 237");
    }

    [TestMethod]
    public void Oct2Hex_Tests()
    {
        this._worksheet.Cells["A1"].Value = "10";
        this._worksheet.Cells["A2"].Formula = "OCT2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("8", this._worksheet.Cells["A2"].Value, "10 was not 8");

        this._worksheet.Cells["A1"].Value = "0000000007";
        this._worksheet.Cells["A2"].Formula = "OCT2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("7", this._worksheet.Cells["A2"].Value, "22 was not 7");

        this._worksheet.Cells["A1"].Value = "10";
        this._worksheet.Cells["A2"].Formula = "OCT2HEX(A1, 10)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("0000000008", this._worksheet.Cells["A2"].Value, "0000000010 was not 0000000008");

        this._worksheet.Cells["A1"].Value = "7777777770";
        this._worksheet.Cells["A2"].Formula = "OCT2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("FFFFFFFFF8", this._worksheet.Cells["A2"].Value, "7777777770 was not FFFFFFFFF8");

        this._worksheet.Cells["A1"].Value = "763";
        this._worksheet.Cells["A2"].Formula = "OCT2HEX(A1)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual("1F3", this._worksheet.Cells["A2"].Value, "763 was not 1F3");
    }

    [TestMethod]
    public void ConvertDistanceTests()
    {
        this._worksheet.Cells["A1"].Value = "1";
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"mi\",\"m\")";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(1609.344d, this._worksheet.Cells["A2"].Value, "1 mile was not 1608 m");

        this._worksheet.Cells["A1"].Value = "4";
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"yd\",\"ft\")";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(12d, this._worksheet.Cells["A2"].Value, "4 yards was not 12 ft");

        this._worksheet.Cells["A1"].Value = "200";
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"cm\",\"m\")";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(2d, this._worksheet.Cells["A2"].Value, "200 cm was not 2 m");
    }

    [TestMethod]
    public void ConvertTimeTests()
    {
        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"yr\",\"day\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(365.25d, result, "1 yr was not 365.25 d");

        this._worksheet.Cells["A1"].Value = 3600;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"sec\",\"hr\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(1d, result, "3600 sec was not 1 hour");
    }

    [TestMethod]
    public void ConvertWeightTests()
    {
        this._worksheet.Cells["A1"].Value = 36;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"kg\",\"g\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(36000d, result, "36 kg was not 36000 g");

        this._worksheet.Cells["A1"].Value = 2;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"sg\",\"lbm\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(64.3481d, result, "2 sg was not 64.3481 lbm");
    }

    [TestMethod]
    public void ConvertSpeedTests()
    {
        this._worksheet.Cells["A1"].Value = 36;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"km/h\",\"m/s\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(10d, result, "36 km/h was not 10 m/s");

        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"admkn\",\"kn\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(1.00064d, result, "36 km/h was not 10 m/s");
    }

    [TestMethod]
    public void ConvertAreaTests()
    {
        this._worksheet.Cells["A1"].Value = 36;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"ha\",\"us_acre\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(88.95758d, result, "16 ha was not 88.95758 us_acre");

        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"km2\",\"m2\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(1000000d, result, "1 km2 was not 1000000 m2");
    }

    [TestMethod]
    public void ConvertLiquidTests()
    {
        this._worksheet.Cells["A1"].Value = 36;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"pt\",\"lt\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(17.03435d, result, "36 pt was not 17.03436 lt");

        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"gal\",\"tsp\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 4);
        Assert.AreEqual(768d, result, "1 gallon was not 768 tsp");

        this._worksheet.Cells["A1"].Value = 1612328564;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"gal\",\"km3\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 4);
        Assert.AreEqual(0.0061, result, "1612328564 gallon was not 0.0061 km3");
    }

    [TestMethod]
    public void ConvertPowerTests()
    {
        this._worksheet.Cells["A1"].Value = 190;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"HP\",\"kw\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(141.682976d, result, "190 horsepowers was not 141.682976 kw");
    }

    [TestMethod]
    public void ConvertMagnetismTests()
    {
        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"T\",\"ga\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(10000d, result, "1 tesla was not 10000 gauss");

        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"ga\",\"T\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(0.0001d, result, "1 gauss was not 0.0001 gauss");
    }

    [TestMethod]
    public void ConvertPressureTests()
    {
        this._worksheet.Cells["A1"].Value = 3;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"mmHg\",\"Torr\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(2.99999, result, "3 mmHg was not 2.99999 Torr");

        this._worksheet.Cells["A1"].Value = 3000;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"p\",\"psi\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(0.435113d, result, "3000 p was not 0.43511 psi");
    }

    [TestMethod]
    public void ConvertTemperatureTests()
    {
        this._worksheet.Cells["A1"].Value = 25;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"C\",\"F\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 5);
        Assert.AreEqual(77d, result, "25 C was not 77 F");

        this._worksheet.Cells["A1"].Value = 25;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"C\",\"kel\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(298.15, result, "25 C was not 298.15 kel");

        this._worksheet.Cells["A1"].Value = 0;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"fah\",\"C\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(-17.777778, result, "0 F was not -17.777778 C");
    }

    [TestMethod]
    public void ConvertEnergyTests()
    {
        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"J\",\"e\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(10000000d, result, "1 J was not 10000000 e");

        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"Me\",\"mWh\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(0.027778d, result, "1 Me was not 0.027778 mWh");
    }

    [TestMethod]
    public void ConvertInformationUnitsTests()
    {
        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"bit\",\"byte\")";
        this._worksheet.Cells["A2"].Calculate();
        double result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(0.125d, result, "1 bit was not 0.125 bytes");

        this._worksheet.Cells["A1"].Value = 2;
        this._worksheet.Cells["A2"].Formula = "CONVERT(A1,\"Gibit\",\"Mibyte\")";
        this._worksheet.Cells["A2"].Calculate();
        result = System.Math.Round((double)this._worksheet.Cells["A2"].Value, 6);
        Assert.AreEqual(256d, result, "2Gbit was not 256 Mbyte");
    }

    [TestMethod]
    public void BitAndTests()
    {
        this._worksheet.Cells["A2"].Formula = "BITAND(5,7)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(5, this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void BitOrTests()
    {
        this._worksheet.Cells["A2"].Formula = "BITOR(9,12)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(13, this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void BitXorTests()
    {
        this._worksheet.Cells["A2"].Formula = "BITXOR(5,6)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(3, this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void BitLshiftTests()
    {
        this._worksheet.Cells["A2"].Formula = "BITLSHIFT(3,5)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(96, this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void BitRshiftTests()
    {
        this._worksheet.Cells["A2"].Formula = "BITRSHIFT(20,2)";
        this._worksheet.Cells["A2"].Calculate();
        Assert.AreEqual(5, this._worksheet.Cells["A2"].Value);
    }
}