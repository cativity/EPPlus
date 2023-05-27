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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.Excel.Functions.RefAndLookup;

[TestClass]
public class LookupNavigatorTests
{
    const string WorksheetName = "";

    private static LookupArguments GetArgs(params object[] args)
    {
        IEnumerable<FunctionArgument>? lArgs = FunctionsHelper.CreateArgs(args);

        return new LookupArguments(lArgs, ParsingContext.Create());
    }

    private static ParsingContext GetContext(ExcelDataProvider provider)
    {
        ParsingContext? ctx = ParsingContext.Create();
        _ = ctx.Scopes.NewScope(new RangeAddress() { Worksheet = WorksheetName, FromCol = 1, FromRow = 1 });
        ctx.ExcelDataProvider = provider;

        return ctx;
    }

    //[TestMethod]
    //public void NavigatorShouldEvaluateFormula()
    //{
    //    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
    //    provider.Stub(x => x.GetCellValue(WorksheetName,0, 0)).Return(new ExcelCell(3);
    //    provider.Stub(x => x.GetCellValue(WorksheetName,1, 0)).Return("B5");
    //    var args = GetArgs(4, "A1:B2", 1);
    //    var context = GetContext(provider);
    //    var parser = MockRepository.GenerateMock<FormulaParser>(provider);
    //    context.Parser = parser;
    //    var navigator = new LookupNavigator(LookupDirection.Vertical, args, context);
    //    navigator.MoveNext();
    //    parser.AssertWasCalled(x => x.Parse("B5"));
    //}

    [TestMethod]
    public void CurrentValueShouldBeFirstCell()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(4);
        LookupArguments? args = GetArgs(3, "A1:B2", 1);
        LookupNavigator? navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
        Assert.AreEqual(3, navigator.CurrentValue);
    }

    [TestMethod]
    public void MoveNextShouldReturnFalseIfLastCell()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(4);
        LookupArguments? args = GetArgs(3, "A1:B1", 1);
        LookupNavigator? navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
        Assert.IsFalse(navigator.MoveNext());
    }

    [TestMethod]
    public void HasNextShouldBeTrueIfNotLastCell()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(5, 5));
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(4);
        LookupArguments? args = GetArgs(3, "A1:B2", 1);
        LookupNavigator? navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
        Assert.IsTrue(navigator.MoveNext());
    }

    [TestMethod]
    public void MoveNextShouldNavigateVertically()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(4);
        _ = A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));
        LookupArguments? args = GetArgs(6, "A1:B2", 1);
        LookupNavigator? navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
        _ = navigator.MoveNext();
        Assert.AreEqual(4, navigator.CurrentValue);
    }

    [TestMethod]
    public void MoveNextShouldIncreaseIndex()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(5, 5));
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(4);
        LookupArguments? args = GetArgs(6, "A1:B2", 1);
        LookupNavigator? navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
        Assert.AreEqual(0, navigator.Index);
        _ = navigator.MoveNext();
        Assert.AreEqual(1, navigator.Index);
    }

    [TestMethod]
    public void GetLookupValueShouldReturnCorrespondingValue()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(5, 5));
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(4);
        LookupArguments? args = GetArgs(6, "A1:B2", 2);
        LookupNavigator? navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
        Assert.AreEqual(4, navigator.GetLookupValue());
    }

    [TestMethod]
    public void GetLookupValueShouldReturnCorrespondingValueWithOffset()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(5, 5));
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
        _ = A.CallTo(() => provider.GetCellValue(WorksheetName, 3, 3)).Returns(4);
        LookupArguments? args = new LookupArguments(3, "A1:A4", 3, 2, false, null);
        LookupNavigator? navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
        Assert.AreEqual(4, navigator.GetLookupValue());
    }
}