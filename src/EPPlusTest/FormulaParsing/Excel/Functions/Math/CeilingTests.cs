using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class CeilingTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(RangeAddress.Empty);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceLowerThan0()
        {
            double expectedValue = 22.36d;
            Ceiling? func = new Ceiling();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(22.35d, 0.01);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingBugTest1()
        {
            double expectedValue = 100d;
            Ceiling? func = new Ceiling();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(100d, 100d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }
        [TestMethod]
        public void CeilingBugTest2()
        {
            double expectedValue = 12000d;
            Ceiling? func = new Ceiling();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(12000d, 1000d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsMinus0point1()
        {
            double expectedValue = -22.4d;
            Ceiling? func = new Ceiling();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-22.35d, -0.1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, System.Math.Round((double)result.Result, 2));
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs1()
        {
            double expectedValue = 23d;
            Ceiling? func = new Ceiling();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(22.35d, 1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs10()
        {
            double expectedValue = 30d;
            Ceiling? func = new Ceiling();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(22.35d, 10);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsNegative()
        {
            double expectedValue = -30d;
            Ceiling? func = new Ceiling();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-22.35d, -10);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldThrowExceptionIfNumberIsPositiveAndSignificanceIsNegative()
        {
            ExcelErrorValue? expectedValue = ExcelErrorValue.Parse("#NUM!");
            Ceiling? func = new Ceiling();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(22.35d, -1);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(expectedValue, result);
        }

        // CEILING.PRECISE
        [TestMethod]
        public void CeilingPreciseShouldHandleSingleArg()
        {
            CeilingPrecise? func = new CeilingPrecise();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6.1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(7d, result.Result);
        }

        [TestMethod]
        public void CeilingMathShouldReturnCorrectResult()
        {
            CeilingMath? func = new CeilingMath();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(15.25);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(16d, result);

            args = FunctionsHelper.CreateArgs(15.25, 0.1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(15.3d, result);

            args = FunctionsHelper.CreateArgs(15.25, 5);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(20d, result);

            args = FunctionsHelper.CreateArgs(-15.25, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-15d, result);

            args = FunctionsHelper.CreateArgs(-15.25, 1, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-16d, result);

            args = FunctionsHelper.CreateArgs(-15.25, 10);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-10d, result);
        }
    }
}
