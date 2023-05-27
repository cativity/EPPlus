using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class FloorTests
{
    private ParsingContext _parsingContext;

    [TestInitialize]
    public void Initialize()
    {
        this._parsingContext = ParsingContext.Create();
        _ = this._parsingContext.Scopes.NewScope(RangeAddress.Empty);
    }

    [TestMethod]
    public void FloorShouldReturnCorrectResultWhenSignificanceIsBetween0And1()
    {
        Floor? func = new Floor();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(26.75d, 0.1);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(26.7d, result.Result);
    }

    [TestMethod]
    public void FloorShouldReturnCorrectResultWhenSignificanceIs1()
    {
        Floor? func = new Floor();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(26.75d, 1);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(26d, result.Result);
    }

    [TestMethod]
    public void FloorShouldReturnCorrectResultWhenSignificanceIsMinus1()
    {
        Floor? func = new Floor();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-26.75d, -1);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(-26d, result.Result);
    }

    [TestMethod]
    public void FloorBugTest1()
    {
        double expectedValue = 100d;
        Floor? func = new Floor();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(100d, 100d);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(expectedValue, result.Result);
    }

    [TestMethod]
    public void FloorBugTest2()
    {
        double expectedValue = 12000d;
        Floor? func = new Floor();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(12000d, 1000d);
        CompileResult? result = func.Execute(args, this._parsingContext);
        Assert.AreEqual(expectedValue, result.Result);
    }

    [TestMethod]
    public void FloorMathShouldReturnCorrectResult()
    {
        FloorMath? func = new FloorMath();

        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(58.55);
        object? result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(58d, result);

        args = FunctionsHelper.CreateArgs(58.55, 0.1);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(58.5d, result);

        args = FunctionsHelper.CreateArgs(58.55, 5);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(55d, result);

        args = FunctionsHelper.CreateArgs(-58.55, 1);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(-59d, result);

        args = FunctionsHelper.CreateArgs(-58.55, 1, 1);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(-58d, result);

        args = FunctionsHelper.CreateArgs(-58.55, 10);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(-60d, result);
    }
}