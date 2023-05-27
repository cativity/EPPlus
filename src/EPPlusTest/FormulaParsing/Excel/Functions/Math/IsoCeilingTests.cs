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

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class IsoCeilingTests
{
    private ParsingContext _parsingContext;

    [TestInitialize]
    public void Initialize()
    {
        this._parsingContext = ParsingContext.Create();
        _ = this._parsingContext.Scopes.NewScope(RangeAddress.Empty);
    }

    [TestMethod]
    public void ShouldReturnCorrectResult()
    {
        IsoCeiling? func = new IsoCeiling();

        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(22.25);
        object? result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(23d, result);

        args = FunctionsHelper.CreateArgs(22.25, 1);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(23d, result);

        args = FunctionsHelper.CreateArgs(22.25, 0.1);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(22.3d, result);

        args = FunctionsHelper.CreateArgs(22.25, 10);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(30d, result);

        args = FunctionsHelper.CreateArgs(-22.25, 1);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(-22d, result);

        args = FunctionsHelper.CreateArgs(-22.25, 0.1);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(-22.2d, result);

        args = FunctionsHelper.CreateArgs(-22.25, 5);
        result = func.Execute(args, this._parsingContext).Result;
        Assert.AreEqual(-20d, result);
    }
}