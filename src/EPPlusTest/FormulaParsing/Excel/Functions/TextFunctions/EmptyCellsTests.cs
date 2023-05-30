using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions;

[TestClass]
public class EmptyCellsTests
{
    private ExcelPackage _package;
    private ExcelWorksheet _worksheet;

    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        this._worksheet = this._package.Workbook.Worksheets.Add("test");
    }

    [TestCleanup]
    public void Cleanup() => this._package.Dispose();

    [TestMethod]
    public void ConcatenateShouldHandleEmptyCells()
    {
        this._worksheet.Cells["A1"].Value = "A";
        this._worksheet.Cells["C1"].Value = "C";
        this._worksheet.Cells["A2"].Formula = "CONCATENATE(A1,B1,C1)";
        this._worksheet.Calculate();
        Assert.AreEqual("AC", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void UpperShouldHandleEmptyCells()
    {
        this._worksheet.Cells["A1"].Value = "A";
        this._worksheet.Cells["C1"].Value = "C";
        this._worksheet.Cells["A2"].Formula = "UPPER(B1)";
        this._worksheet.Calculate();
        Assert.AreEqual("", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void LowerShouldHandleEmptyCells()
    {
        this._worksheet.Cells["A1"].Value = "A";
        this._worksheet.Cells["C1"].Value = "C";
        this._worksheet.Cells["A2"].Formula = "LOWER(B1)";
        this._worksheet.Calculate();
        Assert.AreEqual("", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void ProperHandleEmptyCells()
    {
        this._worksheet.Cells["A1"].Value = "A";
        this._worksheet.Cells["C1"].Value = "C";
        this._worksheet.Cells["A2"].Formula = "PROPER(B1)";
        this._worksheet.Calculate();
        Assert.AreEqual("", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void LeftHandleEmptyCells()
    {
        this._worksheet.Cells["A1"].Value = "A";
        this._worksheet.Cells["C1"].Value = "C";
        this._worksheet.Cells["A2"].Formula = "LEFT(B1,2)";
        this._worksheet.Calculate();
        Assert.AreEqual("", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void RightHandleEmptyCells()
    {
        this._worksheet.Cells["A1"].Value = "A";
        this._worksheet.Cells["C1"].Value = "C";
        this._worksheet.Cells["A2"].Formula = "RIGHT(B1,2)";
        this._worksheet.Calculate();
        Assert.AreEqual("", this._worksheet.Cells["A2"].Value);
    }
}