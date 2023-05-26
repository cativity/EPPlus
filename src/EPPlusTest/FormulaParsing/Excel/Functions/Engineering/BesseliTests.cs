using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering;

[TestClass]
public class BesseliTests
{
    [TestMethod]
    public void BeselliTests()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

        sheet.Cells["A1"].Formula = "BESSELI(4.5, 1)";
        sheet.Calculate();
        object? result = sheet.Cells["A1"].Value;
        Assert.AreEqual(15.389223, System.Math.Round((double)result, 6));
    }

    [TestMethod]
    public void BesellJTests()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

        sheet.Cells["A1"].Formula = "BESSELJ(2.5, 1)";
        sheet.Calculate();
        object? result = sheet.Cells["A1"].Value;
        Assert.AreEqual(0.497094, System.Math.Round((double)result, 6));
    }

    [TestMethod]
    public void BesellKTests()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

        sheet.Cells["A1"].Formula = "BESSELK(0.05, 1)";
        sheet.Calculate();
        object? result = sheet.Cells["A1"].Value;
        Assert.AreEqual(19.909674, System.Math.Round((double)result, 6));
    }

    [TestMethod]
    public void BesellYTests()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

        sheet.Cells["A1"].Formula = "BESSELY(0.05, 1)";
        sheet.Calculate();
        object? result = sheet.Cells["A1"].Value;
        Assert.AreEqual(-12.789855, System.Math.Round((double)result, 6));
    }
}