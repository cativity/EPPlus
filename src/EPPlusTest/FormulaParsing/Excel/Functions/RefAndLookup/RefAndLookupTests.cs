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
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using AddressFunction = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Address;

namespace EPPlusTest.Excel.Functions;

[TestClass]
public class RefAndLookupTests
{
    const string WorksheetName = null;

    [TestMethod]
    public void LookupArgumentsShouldSetSearchedValue()
    {
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, "A:B", 2);
        LookupArguments? lookupArgs = new LookupArguments(args, ParsingContext.Create());
        Assert.AreEqual(1, lookupArgs.SearchedValue);
    }

    [TestMethod]
    public void LookupArgumentsShouldSetRangeAddress()
    {
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, "A:B", 2);
        LookupArguments? lookupArgs = new LookupArguments(args, ParsingContext.Create());
        Assert.AreEqual("A:B", lookupArgs.RangeAddress);
    }

    [TestMethod]
    public void LookupArgumentsShouldSetColIndex()
    {
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, "A:B", 2);
        LookupArguments? lookupArgs = new LookupArguments(args, ParsingContext.Create());
        Assert.AreEqual(2, lookupArgs.LookupIndex);
    }

    [TestMethod]
    public void LookupArgumentsShouldSetRangeLookupToTrueAsDefaultValue()
    {
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, "A:B", 2);
        LookupArguments? lookupArgs = new LookupArguments(args, ParsingContext.Create());
        Assert.IsTrue(lookupArgs.RangeLookup);
    }

    [TestMethod]
    public void LookupArgumentsShouldSetRangeLookupToTrueWhenTrueIsSupplied()
    {
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, "A:B", 2, true);
        LookupArguments? lookupArgs = new LookupArguments(args, ParsingContext.Create());
        Assert.IsTrue(lookupArgs.RangeLookup);
    }

    [TestMethod]
    public void VLookupShouldReturnResultFromMatchingRow()
    {
        VLookup? func = new VLookup();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, "A1:B2", 2);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(2);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);
        _ = A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));

        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(5, result.Result);
    }

    [TestMethod]
    public void VLookupShouldReturnClosestValueBelowWhenRangeLookupIsTrue()
    {
        VLookup? func = new VLookup();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, "A1:B2", 2, true);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(5);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(4);

        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(1, result.Result);
    }

    [TestMethod]
    public void VLookupShouldReturnClosestStringValueBelowWhenRangeLookupIsTrue()
    {
        VLookup? func = new VLookup();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs("B", "A1:B2", 2, true);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        ;

        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns("A");
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns("C");
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(4);

        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(1, result.Result);
    }

    [TestMethod]
    public void HLookupShouldReturnResultFromMatchingRow()
    {
        HLookup? func = new HLookup();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, "A1:B2", 2);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();

        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(2);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(5, result.Result);
    }

    [TestMethod]
    public void HLookupShouldReturnNaErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsFalse()
    {
        HLookup? func = new HLookup();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, "A1:B2", 2, false);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();

        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(2);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        ExcelErrorValue? expectedResult = ExcelErrorValue.Create(eErrorType.NA);
        Assert.AreEqual(expectedResult, result.Result);
    }

    [TestMethod]
    public void HLookupShouldReturnErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsTrue()
    {
        HLookup? func = new HLookup();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, "A1:B2", 2, true);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();

        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(2);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(result.DataType, DataType.ExcelError);
    }

    [TestMethod]
    public void LookupShouldReturnResultFromMatchingRowArrayVertical()
    {
        Lookup? func = new Lookup();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, "A1:B3", 2);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns("A");
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns("B");
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 3, 1)).Returns(5);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 3, 2)).Returns("C");
        _ = A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));

        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual("B", result.Result);
    }

    [TestMethod]
    public void LookupShouldReturnResultFromMatchingRowArrayHorizontal()
    {
        Lookup? func = new Lookup();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, "A1:C2", 2);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 3)).Returns(5);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns("A");
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns("B");
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 3)).Returns("C");

        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual("B", result.Result);
    }

    [TestMethod]
    public void LookupShouldReturnResultFromMatchingSecondArrayHorizontal()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Value = 1;
        sheet.Cells["B1"].Value = 3;
        sheet.Cells["C1"].Value = 5;
        sheet.Cells["A3"].Value = "A";
        sheet.Cells["B3"].Value = "B";
        sheet.Cells["C3"].Value = "C";

        sheet.Cells["D1"].Formula = "LOOKUP(4, A1:C1, A3:C3)";
        sheet.Calculate();
        object? result = sheet.Cells["D1"].Value;
        Assert.AreEqual("B", result);
    }

    [TestMethod]
    public void LookupShouldReturnResultFromMatchingSecondArrayHorizontalWithOffset()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Value = 1;
        sheet.Cells["B1"].Value = 3;
        sheet.Cells["C1"].Value = 5;
        sheet.Cells["B3"].Value = "A";
        sheet.Cells["C3"].Value = "B";
        sheet.Cells["D3"].Value = "C";

        sheet.Cells["D1"].Formula = "LOOKUP(4, A1:C1, B3:D3)";
        sheet.Calculate();
        object? result = sheet.Cells["D1"].Value;
        Assert.AreEqual("B", result);
    }

    [TestMethod]
    public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeExact()
    {
        Match? func = new Match();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(3, "A1:C1", 0);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 3)).Returns(5);
        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(2, result.Result);
    }

    [TestMethod]
    public void MatchShouldReturnIndexOfMatchingValVertical_MatchTypeExact()
    {
        Match? func = new Match();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(3, "A1:A3", 0);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 3, 1)).Returns(5);
        _ = A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));
        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(2, result.Result);
    }

    [TestMethod]
    public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestBelow()
    {
        Match? func = new Match();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, "A1:C1", 1);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 3)).Returns(5);
        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(2, result.Result);
    }

    [TestMethod]
    public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestAbove()
    {
        Match? func = new Match();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6, "A1:C1", -1);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(10);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(8);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 3)).Returns(5);
        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(2, result.Result);
    }

    [TestMethod]
    public void MatchShouldReturnFirstItemWhenExactMatch_MatchTypeClosestAbove()
    {
        Match? func = new Match();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(10, "A1:C1", -1);
        ParsingContext? parsingContext = ParsingContext.Create();
        _ = parsingContext.Scopes.NewScope(RangeAddress.Empty);

        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(10);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(8);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 3)).Returns(5);
        parsingContext.ExcelDataProvider = provider;
        CompileResult? result = func.Execute(args, parsingContext);
        Assert.AreEqual(1, result.Result);
    }

    [TestMethod]
    public void MatchShouldHandleAddressOnOtherSheet()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("Sheet1");
        ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("Sheet2");
        sheet1.Cells["A1"].Formula = "Match(10, Sheet2!A1:Sheet2!A3, 0)";
        sheet2.Cells["A1"].Value = 9;
        sheet2.Cells["A2"].Value = 10;
        sheet2.Cells["A3"].Value = 11;
        sheet1.Calculate();
        Assert.AreEqual(2, sheet1.Cells["A1"].Value);
    }

    [TestMethod]
    public void RowShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
    {
        Row? func = new Row();
        ParsingContext? parsingContext = ParsingContext.Create();
        RangeAddressFactory? rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>());
        _ = parsingContext.Scopes.NewScope(rangeAddressFactory.Create("A2"));
        CompileResult? result = func.Execute(Enumerable.Empty<FunctionArgument>(), parsingContext);
        Assert.AreEqual(2, result.Result);
    }

    [TestMethod]
    public void RowShouldReturnRowSuppliedAddress()
    {
        Row? func = new Row();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("A3"), parsingContext);
        Assert.AreEqual(3, result.Result);
    }

    [TestMethod]
    public void ColumnShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
    {
        Column? func = new Column();
        ParsingContext? parsingContext = ParsingContext.Create();
        RangeAddressFactory? rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>());
        _ = parsingContext.Scopes.NewScope(rangeAddressFactory.Create("B2"));
        CompileResult? result = func.Execute(Enumerable.Empty<FunctionArgument>(), parsingContext);
        Assert.AreEqual(2, result.Result);
    }

    [TestMethod]
    public void ColumnShouldReturnRowSuppliedAddress()
    {
        Column? func = new Column();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("E3"), parsingContext);
        Assert.AreEqual(5, result.Result);
    }

    [TestMethod]
    public void RowsShouldReturnNbrOfRowsSuppliedRange()
    {
        Rows? func = new Rows();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("A1:B3"), parsingContext);
        Assert.AreEqual(3, result.Result);
    }

    [TestMethod]
    public void RowsShouldReturnNbrOfRowsForEntireColumn()
    {
        Rows? func = new Rows();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("A:B"), parsingContext);
        Assert.AreEqual(1048576, result.Result);
    }

    [TestMethod]
    public void ColumnssShouldReturnNbrOfRowsSuppliedRange()
    {
        Columns? func = new Columns();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("A1:E3"), parsingContext);
        Assert.AreEqual(5, result.Result);
    }

    [TestMethod]
    public void ChooseShouldReturnItemByIndex()
    {
        Choose? func = new Choose();
        ParsingContext? parsingContext = ParsingContext.Create();
        CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(1, "A", "B"), parsingContext);
        Assert.AreEqual("A", result.Result);
    }

    [TestMethod]
    public void AddressShouldReturnAddressByIndexWithDefaultRefType()
    {
        Address? func = new AddressFunction();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
        CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(1, 2), parsingContext);
        Assert.AreEqual("$B$1", result.Result);
    }

    [TestMethod]
    public void AddressShouldReturnAddressByIndexWithRelativeType()
    {
        Address? func = new AddressFunction();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
        CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn), parsingContext);
        Assert.AreEqual("B1", result.Result);
    }

    [TestMethod]
    public void AddressShouldReturnAddressByWithSpecifiedWorksheet()
    {
        Address? func = new AddressFunction();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);

        CompileResult? result =
            func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, true, "Worksheet1"), parsingContext);

        Assert.AreEqual("Worksheet1!B1", result.Result);
    }

    [TestMethod, ExpectedException(typeof(InvalidOperationException))]
    public void AddressShouldThrowIfR1C1FormatIsSpecified()
    {
        Address? func = new AddressFunction();
        ParsingContext? parsingContext = ParsingContext.Create();
        parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
        _ = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, false), parsingContext);
    }
}