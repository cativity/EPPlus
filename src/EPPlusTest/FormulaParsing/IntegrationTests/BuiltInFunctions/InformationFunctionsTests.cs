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
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public class InformationFunctionsTests : FormulaParserTestBase
{
    private ExcelPackage _package;

    [TestInitialize]
    public void Setup()
    {
        this._package = new ExcelPackage();
        _ = this._package.Workbook.Worksheets.Add("test");
        this._parser = new FormulaParser(this._package);
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._package.Dispose();
        this._package = null;
    }

    [TestMethod]
    public void IsBlankShouldReturnCorrectValue()
    {
        object? result = this._parser.Parse("ISBLANK(A1)");
        Assert.IsTrue((bool)result);
    }

    [TestMethod]
    public void IsNumberShouldReturnCorrectValue()
    {
        object? result = this._parser.Parse("ISNUMBER(10/2)");
        Assert.IsTrue((bool)result);
    }

    [TestMethod]
    public void IsErrorShouldReturnTrueWhenDivBy0()
    {
        object? result = this._parser.Parse("ISERROR(10/0)");
        Assert.IsTrue((bool)result);
    }

    [TestMethod]
    public void IsTextShouldReturnTrueWhenReferencedCellContainsText()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("Test");
        sheet.Cells["A1"].Value = "Abc";
        sheet.Cells["A2"].Formula = "ISTEXT(A1)";
        sheet.Calculate();
        object? result = sheet.Cells["A2"].Value;
        Assert.IsTrue((bool)result);
    }

    [TestMethod]
    public void IsErrShouldReturnFalseIfErrorCodeIsNa()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("Test");
        sheet.Cells["A1"].Value = ExcelErrorValue.Parse("#N/A");
        sheet.Cells["A2"].Formula = "ISERR(A1)";
        sheet.Calculate();
        object? result = sheet.Cells["A2"].Value;
        Assert.IsFalse((bool)result);
    }

    [TestMethod]
    public void IsNaShouldReturnTrueCodeIsNa()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("Test");
        sheet.Cells["A1"].Value = ExcelErrorValue.Parse("#N/A");
        sheet.Cells["A2"].Formula = "ISNA(A1)";
        sheet.Calculate();
        object? result = sheet.Cells["A2"].Value;
        Assert.IsTrue((bool)result);
    }

    [TestMethod]
    public void ErrorTypeShouldReturnCorrectErrorCodes()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("Test");
        sheet.Cells["A1"].Value = ExcelErrorValue.Create(eErrorType.Null);
        sheet.Cells["B1"].Formula = "ERROR.TYPE(A1)";
        sheet.Cells["A2"].Value = ExcelErrorValue.Create(eErrorType.Div0);
        sheet.Cells["B2"].Formula = "ERROR.TYPE(A2)";
        sheet.Cells["A3"].Value = ExcelErrorValue.Create(eErrorType.Value);
        sheet.Cells["B3"].Formula = "ERROR.TYPE(A3)";
        sheet.Cells["A4"].Value = ExcelErrorValue.Create(eErrorType.Ref);
        sheet.Cells["B4"].Formula = "ERROR.TYPE(A4)";
        sheet.Cells["A5"].Value = ExcelErrorValue.Create(eErrorType.Name);
        sheet.Cells["B5"].Formula = "ERROR.TYPE(A5)";
        sheet.Cells["A6"].Value = ExcelErrorValue.Create(eErrorType.Num);
        sheet.Cells["B6"].Formula = "ERROR.TYPE(A6)";
        sheet.Cells["A7"].Value = ExcelErrorValue.Create(eErrorType.NA);
        sheet.Cells["B7"].Formula = "ERROR.TYPE(A7)";
        sheet.Cells["A8"].Value = 10;
        sheet.Cells["B8"].Formula = "ERROR.TYPE(A8)";
        sheet.Calculate();
        object? nullResult = sheet.Cells["B1"].Value;
        object? div0Result = sheet.Cells["B2"].Value;
        object? valueResult = sheet.Cells["B3"].Value;
        object? refResult = sheet.Cells["B4"].Value;
        object? nameResult = sheet.Cells["B5"].Value;
        object? numResult = sheet.Cells["B6"].Value;
        object? naResult = sheet.Cells["B7"].Value;
        object? noErrorResult = sheet.Cells["B8"].Value;
        Assert.AreEqual(1, nullResult, "Null error was not 1");
        Assert.AreEqual(2, div0Result, "Div0 error was not 2");
        Assert.AreEqual(3, valueResult, "Value error was not 3");
        Assert.AreEqual(4, refResult, "Ref error was not 4");
        Assert.AreEqual(5, nameResult, "Name error was not 5");
        Assert.AreEqual(6, numResult, "Num error was not 6");
        Assert.AreEqual(7, naResult, "NA error was not 7");
        Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), noErrorResult, "No error did not return N/A error");
    }
}