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
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public class RefAndLookupTests : FormulaParserTestBase
{
    private ExcelDataProvider _excelDataProvider;
    const string WorksheetName = null;
    private ExcelPackage _package;
    private ExcelWorksheet _worksheet;

    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        this._worksheet = this._package.Workbook.Worksheets.Add("Test");
        this._excelDataProvider = A.Fake<ExcelDataProvider>();
        A.CallTo(() => this._excelDataProvider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(10, 1));
        A.CallTo(() => this._excelDataProvider.GetWorkbookNameValues()).Returns(new ExcelNamedRangeCollection(this._package.Workbook));
        this._parser = new FormulaParser(this._excelDataProvider);
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._package.Dispose();
    }

    [TestMethod]
    public void VLookupShouldReturnCorrespondingValue()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("test");
        string? lookupAddress = "A1:B2";
        ws.Cells["A1"].Value = 1;
        ws.Cells["B1"].Value = 1;
        ws.Cells["A2"].Value = 2;
        ws.Cells["B2"].Value = 5;
        ws.Cells["A3"].Formula = "VLOOKUP(2, " + lookupAddress + ", 2)";
        ws.Calculate();
        object? result = ws.Cells["A3"].Value;
        Assert.AreEqual(5, result);
    }

    [TestMethod]
    public void VLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("test");
        string? lookupAddress = "A1:B2";
        ws.Cells["A1"].Value = 3;
        ws.Cells["B1"].Value = 1;
        ws.Cells["A2"].Value = 5;
        ws.Cells["B2"].Value = 5;
        ws.Cells["A3"].Formula = "VLOOKUP(4, " + lookupAddress + ", 2, true)";
        ws.Calculate();
        object? result = ws.Cells["A3"].Value;
        Assert.AreEqual(1, result);
    }

    [TestMethod]
    public void HLookupShouldReturnCorrespondingValue()
    {
        string? lookupAddress = "A1:B2";
        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["B1"].Value = 2;
        this._worksheet.Cells["A2"].Value = 2;
        this._worksheet.Cells["B2"].Value = 5;
        this._worksheet.Cells["A3"].Formula = "HLOOKUP(2, " + lookupAddress + ", 2)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A3"].Value;
        Assert.AreEqual(5, result);
    }

    [TestMethod]
    public void HLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
    {
        string? lookupAddress = "A1:B2";
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s = package.Workbook.Worksheets.Add("test");
        s.Cells[1, 1].Value = 3;
        s.Cells[1, 2].Value = 5;
        s.Cells[2, 1].Value = 1;
        s.Cells[2, 2].Value = 2;
        s.Cells[5, 5].Formula = "HLOOKUP(4, " + lookupAddress + ", 2, true)";
        s.Calculate();
        Assert.AreEqual(1, s.Cells[5, 5].Value);
    }

    [TestMethod]
    public void LookupShouldReturnMatchingValue()
    {
        string? lookupAddress = "A1:B2";
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s = package.Workbook.Worksheets.Add("test");
        s.Cells[1, 1].Value = 3;
        s.Cells[1, 2].Value = 5;
        s.Cells[2, 1].Value = 4;
        s.Cells[2, 2].Value = 1;
        s.Cells[5, 5].Formula = "LOOKUP(4, " + lookupAddress + ")";
        s.Calculate();
        Assert.AreEqual(1, s.Cells[5, 5].Value);
        //    A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,1, 1)).Returns(3);
        //A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,1, 2)).Returns(5);
        //A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,2, 1)).Returns(4);
        //A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,2, 2)).Returns(1);
        //var result = _parser.Parse("LOOKUP(4, " + lookupAddress + ")");
        //Assert.AreEqual(1, result);
    }

    [TestMethod]
    public void MatchShouldReturnIndexOfMatchingValue()
    {
        string? lookupAddress = "A1:A2";

        this._worksheet.Cells["A1"].Value = 3;
        this._worksheet.Cells["A2"].Value = 5;
        this._worksheet.Cells["A3"].Formula = "MATCH(3, " + lookupAddress + ")";
        this._worksheet.Calculate();
        Assert.AreEqual(1, this._worksheet.Cells["A3"].Value);

    }

    [TestMethod]
    public void RowShouldReturnRowNumber()
    {
        A.CallTo(() => this._excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Row()");
        object? result = this._parser.ParseAt("A4");
        Assert.AreEqual(4, result);
    }

    [TestMethod]
    public void RowSholdHandleReference()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("test");
        s1.Cells["A1"].Formula = "ROW(A4)";
        s1.Calculate();
        Assert.AreEqual(4, s1.Cells["A1"].Value);
    }

    [TestMethod]
    public void ColumnShouldReturnRowNumber()
    {
        A.CallTo(() => this._excelDataProvider.GetRangeFormula("", 4, 2)).Returns("Column()");
        object? result = this._parser.ParseAt("B4");
        Assert.AreEqual(2, result);
    }

    [TestMethod]
    public void ColumnSholdHandleReference()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("test");
        s1.Cells["A1"].Formula = "COLUMN(B4)";
        s1.Calculate();
        Assert.AreEqual(2, s1.Cells["A1"].Value);
    }

    [TestMethod]
    public void RowsShouldReturnNbrOfRows()
    {
        A.CallTo(() => this._excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Rows(A5:B7)");
        A.CallTo(() => this._excelDataProvider.GetRange("", 4, 1, "A5:B7")).Returns(new EpplusExcelDataProvider.RangeInfo(this._worksheet, 1, 2, 3, 3));
        object? result = this._parser.ParseAt("A4");
        Assert.AreEqual(3, result);
    }

    [TestMethod]
    public void ColumnsShouldReturnNbrOfCols()
    {
        A.CallTo(() => this._excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Columns(A5:B7)");
        A.CallTo(() => this._excelDataProvider.GetRange("", 4, 1, "A5:B7")).Returns(new EpplusExcelDataProvider.RangeInfo(this._worksheet, 1, 2, 1, 3));
        object? result = this._parser.ParseAt("A4");
        Assert.AreEqual(2, result);
    }

    [TestMethod]
    public void ChooseShouldReturnCorrectResult()
    {
        object? result = this._parser.Parse("Choose(1, \"A\", \"B\")");
        Assert.AreEqual("A", result);
    }

    [TestMethod]
    public void AddressShouldReturnCorrectResult()
    {
        A.CallTo(() => this._excelDataProvider.ExcelMaxRows).Returns(12345);
        object? result = this._parser.Parse("Address(1, 1)");
        Assert.AreEqual("$A$1", result);
    }

    [TestMethod]
    public void IndirectShouldReturnARange()
    {
        using ExcelPackage? package = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("Test");
        s1.Cells["A1:A2"].Value = 2;
        s1.Cells["A3"].Formula = "SUM(Indirect(\"A1:A2\"))";
        s1.Calculate();
        Assert.AreEqual(4d, s1.Cells["A3"].Value);

        s1.Cells["A4"].Formula = "SUM(Indirect(\"A1:A\" & \"2\"))";
        s1.Calculate();
        Assert.AreEqual(4d, s1.Cells["A4"].Value);
    }

    [TestMethod]
    public void OffsetShouldReturnASingleValue()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("Test");
        s1.Cells["B3"].Value = 1d;
        s1.Cells["A5"].Formula = "OFFSET(A1, 2, 1)";
        s1.Calculate();
        Assert.AreEqual(1d, s1.Cells["A5"].Value);
    }

    [TestMethod]
    public void OffsetShouldReturnARange()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("Test");
        s1.Cells["B1"].Value = 1d;
        s1.Cells["B2"].Value = 1d;
        s1.Cells["B3"].Value = 1d;
        s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1))";
        s1.Calculate();
        Assert.AreEqual(3d, s1.Cells["A5"].Value);
    }

    [TestMethod]
    public void OffsetShouldReturnARange2()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("Test");
        s1.Cells["B1"].Value = 10d;
        s1.Cells["B2"].Value = 10d;
        s1.Cells["B3"].Value = 10d;
        s1.Cells["A5"].Formula = "COUNTA(OFFSET(Test!B1, 0, 0, Test!B2, Test!B3))";
        s1.Calculate();
        Assert.AreEqual(3d, s1.Cells["A5"].Value);
    }

    [TestMethod]
    public void OffsetDirectReferenceToMultiRangeShouldSetValueError()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("Test");
        s1.Cells["B1"].Value = 1d;
        s1.Cells["B2"].Value = 1d;
        s1.Cells["B3"].Value = 1d;
        s1.Cells["A5"].Formula = "OFFSET(A1:A3, 0, 1)";
        s1.Calculate();
        object? result = s1.Cells["A5"].Value;
        Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
    }

    [TestMethod]
    public void OffsetShouldReturnARangeAccordingToWidth()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("Test");
        s1.Cells["B1"].Value = 1d;
        s1.Cells["B2"].Value = 1d;
        s1.Cells["B3"].Value = 1d;
        s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2))";
        s1.Calculate();
        Assert.AreEqual(2d, s1.Cells["A5"].Value);
    }

    [TestMethod]
    public void OffsetShouldReturnARangeAccordingToHeight()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("Test");
        s1.Cells["B1"].Value = 1d;
        s1.Cells["B2"].Value = 1d;
        s1.Cells["B3"].Value = 1d;
        s1.Cells["C1"].Value = 2d;
        s1.Cells["C2"].Value = 2d;
        s1.Cells["C3"].Value = 2d;
        s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2, 2))";
        s1.Calculate();
        Assert.AreEqual(6d, s1.Cells["A5"].Value);
    }

    [TestMethod]
    public void OffsetShouldCoverMultipleColumns()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s1 = package.Workbook.Worksheets.Add("Test");
        s1.Cells["C1"].Value = 1d;
        s1.Cells["C2"].Value = 1d;
        s1.Cells["C3"].Value = 1d;
        s1.Cells["D1"].Value = 2d;
        s1.Cells["D2"].Value = 2d;
        s1.Cells["D3"].Value = 2d;
        s1.Cells["A5"].Formula = "SUM(OFFSET(A1:B3, 0, 2))";
        s1.Calculate();
        Assert.AreEqual(9d, s1.Cells["A5"].Value);
    }

    [TestMethod, Ignore]
    public void VLookupShouldHandleNames()
    {
        using ExcelPackage? package = new ExcelPackage(new FileInfo(@"c:\temp\Book3.xlsx"));
        ExcelWorksheet? s1 = package.Workbook.Worksheets.First();
        string? v = s1.Cells["X10"].Formula;
        //s1.Calculate();
        v = s1.Cells["X10"].Formula;
    }

    [TestMethod]
    public void LookupShouldReturnFromResultVector()
    {
        string? lookupAddress = "A1:A5";
        string? resultAddress = "B1:B5";
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s = package.Workbook.Worksheets.Add("test");
        //lookup_vector
        s.Cells[1, 1].Value = 4.14;
        s.Cells[2, 1].Value = 4.19;
        s.Cells[3, 1].Value = 5.17;
        s.Cells[4, 1].Value = 5.77;
        s.Cells[5, 1].Value = 6.39;
        //result_vector
        s.Cells[1, 2].Value = "red";
        s.Cells[2, 2].Value = "orange";
        s.Cells[3, 2].Value = "yellow";
        s.Cells[4, 2].Value = "green";
        s.Cells[5, 2].Value = "blue";
        //lookup_value
        s.Cells[1, 3].Value = 4.14;
        s.Cells[5, 5].Formula = "LOOKUP(C1, " + lookupAddress + ", " + resultAddress + ")";
        s.Calculate();
        Assert.AreEqual("red", s.Cells[5, 5].Value);
    }

    [TestMethod]
    public void LookupShouldCompareEqualDateWithDouble()
    {
        DateTime date = new DateTime(2020, 2, 7).Date;
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s = package.Workbook.Worksheets.Add("test");
        //lookup_vector
        s.Cells[1, 1].Value = date;
        //result vector
        s.Cells[1, 2].Value = 10;

        //lookup value
        s.Cells[1, 3].Value = date;
        s.Cells[1, 4].Formula = "LOOKUP(C1, A1:A2, B1:B2)";
        s.Calculate();
        Assert.AreEqual(10, s.Cells[1, 4].Value);
    }
    [TestMethod]
    public void OffsetInSecondPartOfRange()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s = package.Workbook.Worksheets.Add("test");
        package.Workbook.FormulaParser.Configure(x => x.AllowCircularReferences = true);
        s.Cells[1, 1].Value = 3;
        s.Cells[2, 1].Value = 5;
        s.Cells[3, 1].Formula = "SUM(A1:OFFSET(A3,-1,0))";
        s.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });
        Assert.AreEqual(8d, s.Cells[3, 1].Value);
    }

    [TestMethod]
    public void OffsetInFirstPartOfRange()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s = package.Workbook.Worksheets.Add("test");
        s.Cells[1, 1].Value = 3;
        s.Cells[2, 1].Value = 5;
        s.Cells[4, 1].Formula = "SUM(OFFSET(A3,-1,0):A1)";
        s.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });
        Assert.AreEqual(8d, s.Cells[4, 1].Value);
    }

    [TestMethod]
    public void OffsetInBothPartsOfRange()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? s = package.Workbook.Worksheets.Add("test");
        s.Cells[1, 1].Value = 3;
        s.Cells[2, 1].Value = 5;
        s.Cells[4, 1].Formula = "SUM(OFFSET(A3,-2,0):OFFSET(A3,-1,0))";
        s.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });
        Assert.AreEqual(8d, s.Cells[4, 1].Value);
    }
}