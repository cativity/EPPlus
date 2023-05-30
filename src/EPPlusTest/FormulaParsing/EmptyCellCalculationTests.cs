using EPPlusTest.FormulaParsing.IntegrationTests;
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing;

[TestClass]
public class EmptyCellCalculationTests
{
    private ExcelWorksheet _sheet;
    private ExcelPackage _package;
    //private FormulaParser _parser;

    [TestInitialize]
    public void Setup()
    {
        this._package = new ExcelPackage();
        this._sheet = this._package.Workbook.Worksheets.Add("Test");
        //this._parser = this._package.Workbook.FormulaParser;
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._package.Dispose();
        //this._parser = null;
    }

    [TestMethod]
    public void EmptyCellReferenceShouldReturnZero()
    {
        this._sheet.Cells["A2"].Formula = "A1";
        this._sheet.Calculate();
        object? result = this._sheet.Cells["A2"].Value;
        Assert.AreEqual(0d, result);
    }

    [TestMethod]
    public void EmptyCellReferenceMultiplicationShouldReturnZero()
    {
        this._sheet.Cells["A2"].Formula = "A1*3";
        this._sheet.Calculate();
        object? result = this._sheet.Cells["A2"].Value;
        Assert.AreEqual(0d, result);
    }

    [TestMethod]
    public void EmptyCellReferenceAdditionShouldReturnOtherOperand()
    {
        this._sheet.Cells["A2"].Formula = "A1+2";
        this._sheet.Calculate();
        object? result = this._sheet.Cells["A2"].Value;
        Assert.AreEqual(2d, result);
    }

    [TestMethod]
    public void IfResultEmptyCellReferenceReturnsZero()
    {
        this._sheet.Cells["A1"].Formula = "IF(TRUE,A2)";
        this._sheet.Calculate();
        object? result = this._sheet.Cells["A1"].Value;
        Assert.AreEqual(0d, result);
    }

    [TestMethod]
    public void EmptyCellReferenceShouldEqualZero()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("Test");
        sheet.Cells["A2"].Formula = "A1=0";
        sheet.Calculate();
        Assert.IsTrue((bool)sheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void EmptyCellReferenceShouldEqualEmptyString()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("Test");
        sheet.Cells["A2"].Formula = "A1=\"\"";
        sheet.Calculate();
        Assert.IsTrue((bool)sheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void EmptyCellReferenceShouldEqualFalse()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("Test");
        sheet.Cells["A2"].Formula = "A1=FALSE";
        sheet.Calculate();
        Assert.IsTrue((bool)sheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void IfConditionEmptyCellReferenceEqualsZero()
    {
        this._sheet.Cells["A2"].Formula = "IF(A1=0,1)";
        this._sheet.Calculate();
        object? result = this._sheet.Cells["A2"].Value;
        Assert.AreEqual(1d, result);
    }

    [TestMethod]
    public void IfConditionEmptyCellReferenceEqualsEmptyString()
    {
        this._sheet.Cells["A2"].Formula = "IF(A1=\"\",1)";
        this._sheet.Calculate();
        object? result = this._sheet.Cells["A2"].Value;
        Assert.AreEqual(1d, result);
    }

    [TestMethod]
    public void IfConditionEmptyCellReferenceEqualsFalse()
    {
        this._sheet.Cells["A2"].Formula = "IF(A1=FALSE,1)";
        this._sheet.Calculate();
        object? result = this._sheet.Cells["A2"].Value;
        Assert.AreEqual(1d, result);
    }
}