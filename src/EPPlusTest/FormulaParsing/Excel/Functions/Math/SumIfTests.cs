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

using System.Collections.Generic;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class SumIfTests
{
    private ExcelPackage _package;
    private EpplusExcelDataProvider _provider;
    private ParsingContext _parsingContext;
    private ExcelWorksheet _worksheet;

    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        this._provider = new EpplusExcelDataProvider(this._package);
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
    public void SumIfNumeric()
    {
        this._worksheet.Cells["A1"].Value = 1d;
        this._worksheet.Cells["A2"].Value = 2d;
        this._worksheet.Cells["A3"].Value = 3d;
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">1", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(8d, result.Result);
    }

    [TestMethod]
    public void SumIfNonNumeric()
    {
        this._worksheet.Cells["A1"].Value = "Monday";
        this._worksheet.Cells["A2"].Value = "Tuesday";
        this._worksheet.Cells["A3"].Value = "Thursday";
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "T*day", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(8d, result.Result);
    }

    [TestMethod]
    public void SumIfShouldIgnoreNumericStrings()
    {
        this._worksheet.Cells["A1"].Value = 2;
        this._worksheet.Cells["A2"].Value = 1;
        this._worksheet.Cells["A3"].Value = "4";
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">1");
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(2d, result.Result);
    }

    [TestMethod]
    public void SumIfNumericExpression()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = 1d;
        this._worksheet.Cells["A3"].Value = "Not Empty";
        SumIf? func = new SumIf();
        IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, 1d);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(1d, result.Result);
    }

    [TestMethod]
    public void SumIfEqualToEmptyString()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = "Not Empty";
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(4d, result.Result);
    }

    [TestMethod]
    public void SumIfNotEqualToNull()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = "Not Empty";
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<>", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(8d, result.Result);
    }

    [TestMethod]
    public void SumIfEqualToZero()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = 0d;
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "0", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void SumIfNotEqualToZero()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = 0d;
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<>0", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(4d, result.Result);
    }

    [TestMethod]
    public void SumIfGreaterThanZero()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = 1d;
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">0", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void SumIfGreaterThanOrEqualToZero()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = 1d;
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">=0", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void SumIfLessThanZero()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = -1d;
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<0", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void SumIfLessThanOrEqualToZero()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = -1d;
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<=0", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void SumIfLessThanCharacter()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = "Not Empty";
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<a", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(3d, result.Result);
    }

    [TestMethod]
    public void SumIfLessThanOrEqualToCharacter()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = "Not Empty";
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<=a", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(3d, result.Result);
    }

    [TestMethod]
    public void SumIfGreaterThanCharacter()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = "Not Empty";
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">a", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void SumIfGreaterThanOrEqualToCharacter()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = "Not Empty";
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void SumIfHandleDates()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = string.Empty;
        this._worksheet.Cells["A3"].Value = "Not Empty";
        this._worksheet.Cells["B1"].Value = 1d;
        this._worksheet.Cells["B2"].Value = 3d;
        this._worksheet.Cells["B3"].Value = 5d;
        SumIf? func = new SumIf();
        IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
        IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void SumIfShouldHandleBooleanArg()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Value = true;
        sheet.Cells["B1"].Value = 1;
        sheet.Cells["A2"].Value = false;
        sheet.Cells["B2"].Value = 1;
        sheet.Cells["C1"].Formula = "SUMIF(A1:A2,TRUE,B1:B2)";
        sheet.Calculate();
        Assert.AreEqual(1d, sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void SumIfShouldHandleArrayOfCriterias()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Value = "A";
        sheet.Cells["A2"].Value = "B";
        sheet.Cells["A3"].Value = "A";
        sheet.Cells["A4"].Value = "B";
        sheet.Cells["A5"].Value = "C";
        sheet.Cells["A6"].Value = "B";

        sheet.Cells["B1"].Value = 10;
        sheet.Cells["B2"].Value = 20;
        sheet.Cells["B3"].Value = 10;
        sheet.Cells["B4"].Value = 30;
        sheet.Cells["B5"].Value = 40;
        sheet.Cells["B6"].Value = 10;

        sheet.Cells["A9"].Formula = "SUMIF(A1:A6,{\"A\",\"C\"}, B1:B6)";
        sheet.Calculate();
        Assert.AreEqual(60d, sheet.Cells["A9"].Value);
    }

    [TestMethod]
    public void SumIfShouldHandleRangeWithCriterias()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Value = "A";
        sheet.Cells["A2"].Value = "B";
        sheet.Cells["A3"].Value = "A";
        sheet.Cells["A4"].Value = "B";
        sheet.Cells["A5"].Value = "C";
        sheet.Cells["A6"].Value = "B";

        sheet.Cells["B1"].Value = 10;
        sheet.Cells["B2"].Value = 20;
        sheet.Cells["B3"].Value = 10;
        sheet.Cells["B4"].Value = 30;
        sheet.Cells["B5"].Value = 40;
        sheet.Cells["B6"].Value = 10;

        sheet.Cells["D9"].Value = "A";
        sheet.Cells["E9"].Value = "C";

        sheet.Cells["A9"].Formula = "SUMIF(A1:A6,D9:E9, B1:B6)";
        sheet.Calculate();
        Assert.AreEqual(60d, sheet.Cells["A9"].Value);
    }

    [TestMethod]
    public void SumIfSingleCell()
    {
        this._worksheet.Cells["A1"].Value = 20;
        this._worksheet.Cells["A2"].Formula = "SUMIF(A1,\">0\")";
        this._worksheet.Calculate();

        Assert.AreEqual(20d, this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void SumIfEqualToEmptyString_Parser()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Value = null;
        sheet.Cells["A2"].Value = string.Empty;
        sheet.Cells["A3"].Value = "Not Empty";
        sheet.Cells["B1"].Value = 1d;
        sheet.Cells["B2"].Value = 3d;
        sheet.Cells["B3"].Value = 5d;

        sheet.Cells["B4"].Formula = "SUMIF(A1:A3,\"=\",B1:B3)";
        sheet.Cells["B5"].Formula = "SUMIF(A1:A3,\"\",B1:B3)";
        sheet.Calculate();
        Assert.AreEqual(1d, sheet.Cells["B4"].Value);
        Assert.AreEqual(4d, sheet.Cells["B5"].Value);
    }

    [TestMethod]
    public void SumIfNotEqualToNull_Parser()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Value = null;
        sheet.Cells["A2"].Value = string.Empty;
        sheet.Cells["A3"].Value = "Not Empty";
        sheet.Cells["B1"].Value = 1d;
        sheet.Cells["B2"].Value = 3d;
        sheet.Cells["B3"].Value = 5d;

        sheet.Cells["B4"].Formula = "SUMIF(A1:A3,\"<>\",B1:B3)";
        sheet.Calculate();
        Assert.AreEqual(8d, sheet.Cells["B4"].Value);
    }
}