﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing;

[TestClass]
public class NumericFormulaTests
{
    [TestMethod]
    public void ShouldHandleNumericFormulaLow()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells[1, 1].Formula = "4.00801603206413E-06";
        sheet.Calculate();
        Assert.AreEqual(4.00801603206413E-06d, sheet.Cells[1, 1].Value);
    }

    [TestMethod]
    public void ShouldHandleNegativeNumericFormulaLow()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells[1, 1].Formula = "-4.00801603206413E-06";
        sheet.Calculate();
        Assert.AreEqual(-4.00801603206413E-06d, sheet.Cells[1, 1].Value);
    }

    public static void ShouldHandleNumericFormulaHigh()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells[1, 1].Formula = "4.00801603206413E+06";
        sheet.Calculate();
        Assert.AreEqual(4.00801603206413E+06d, sheet.Cells[1, 1].Value);
    }

    [TestMethod]
    public void ShouldHandleNegativeNumericFormulaHigh()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells[1, 1].Formula = "-4.00801603206413E+06";
        sheet.Calculate();
        Assert.AreEqual(-4.00801603206413E+06d, sheet.Cells[1, 1].Value);
    }

    [TestMethod]
    public void ShouldHandleNumericFormulaWithOperator()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells[1, 1].Formula = "4.00801603206413E-06 + 1";
        sheet.Calculate();
        double result = Math.Round((double)sheet.Cells[1, 1].Value, 9);
        Assert.AreEqual(1.000004008, result);
    }

    [TestMethod]
    public void ShouldHandleIntegerWithScientificNotation()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells[1, 1].Formula = "1E+30";
        sheet.Calculate();
        double result = Math.Round((double)sheet.Cells[1, 1].Value, 9);
        Assert.AreEqual(1E+30, result);
    }

    [TestMethod]
    public void ShouldHandleIntegerWithScientificNotation_IgnoreWhitespce()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Value = "Active";
        sheet.Cells["B1"].Formula = "IF(A1 = \"Active\", 9E+30, 99)";
        sheet.Calculate();
        object? v = sheet.Cells["B1"].Value;
        Assert.AreEqual(9E+30, v);
    }
}