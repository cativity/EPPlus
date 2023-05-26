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
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class MathFunctionsTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(RangeAddress.Empty);
        }

        [TestMethod]
        public void PiShouldReturnPIConstant()
        {
            double expectedValue = (double)Math.Round(Math.PI, 14);
            Pi? func = new Pi();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(0);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void AbsShouldReturnCorrectResult()
        {
            double expectedValue = 3d;
            Abs? func = new Abs();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-3d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void AsinShouldReturnCorrectResult()
        {
            const double expectedValue = 1.5708;
            Asin? func = new Asin();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1d);
            CompileResult? result = func.Execute(args, _parsingContext);
            double rounded = Math.Round((double)result.Result, 4);
            Assert.AreEqual(expectedValue, rounded);
        }

        [TestMethod]
        public void AsinhShouldReturnCorrectResult()
        {
            const double expectedValue = 0.0998;
            Asinh? func = new Asinh();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(0.1d);
            CompileResult? result = func.Execute(args, _parsingContext);
            double rounded = Math.Round((double)result.Result, 4);
            Assert.AreEqual(expectedValue, rounded);
        }

        [TestMethod]
        public void CombinShouldReturnCorrectResult_6_2()
        {
            Combin? func = new Combin();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(15d, result.Result);
        }

        [TestMethod]
        public void CombinShouldReturnCorrectResult_decimal()
        {
            Combin? func = new Combin();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(10.456, 3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(120d, result.Result);
        }

        [TestMethod]
        public void CombinShouldReturnCorrectResult_6_1()
        {
            Combin? func = new Combin();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6, 1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(6d, result.Result);
        }

        [TestMethod]
        public void CombinaShouldReturnCorrectResult_6_2()
        {
            Combina? func = new Combina();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(21d, result.Result);
        }

        [TestMethod]
        public void CombinaShouldReturnCorrectResult_6_5()
        {
            Combina? func = new Combina();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6, 5);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(252d, result.Result);
        }

        [TestMethod]
        public void PermutationaShouldReturnCorrectResult()
        {
            Permutationa? func = new Permutationa();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6, 6);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(46656d, result.Result);

            args = FunctionsHelper.CreateArgs(10, 6);
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1000000d, result.Result);
        }

        [TestMethod]
        public void PermutShouldReturnCorrectResult()
        {
            Permut? func = new Permut();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6, 6);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(720d, result.Result);

            args = FunctionsHelper.CreateArgs(10, 6);
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(151200d, result.Result);
        }

        [TestMethod]
        public void SecShouldReturnCorrectResult_MinusPi()
        {
            Sec? func = new Sec();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-3.14159265358979);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-1d, result.Result);
        }

        [TestMethod]
        public void SecShouldReturnCorrectResult_Zero()
        {
            Sec? func = new Sec();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(0d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SechShouldReturnCorrectResult_PiDividedBy4()
        {
            SecH? func = new SecH();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(Math.PI / 4);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(0.7549, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void SechShouldReturnCorrectResult_MinusPi()
        {
            SecH? func = new SecH();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-3.14159265358979);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(0.08627, Math.Round((double)result, 5));
        }

        [TestMethod]
        public void SechShouldReturnCorrectResult_Zero()
        {
            SecH? func = new SecH();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(0d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SecShouldReturnCorrectResult_PiDividedBy4()
        {
            Sec? func = new Sec();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(Math.PI / 4);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(1.4142, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void CscShouldReturnCorrectResult_Minus6()
        {
            Csc? func = new Csc();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-6);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(3.5789, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void CotShouldReturnCorrectResult_2()
        {
            Cot? func = new Cot();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-0.4577, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void CothShouldReturnCorrectResult_MinusPi()
        {
            Coth? func = new Coth();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(Math.PI * -1);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-1.0037, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void AcothShouldReturnCorrectResult_MinusPi()
        {
            Acoth? func = new Acoth();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-5);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-0.2027, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void RadiansShouldReturnCorrectResult_50()
        {
            Radians? func = new Radians();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(50);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(0.8727, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void RadiansShouldReturnCorrectResult_360()
        {
            Radians? func = new Radians();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(360);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(6.2832, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void AcotShouldReturnCorrectResult_1()
        {
            Acot? func = new Acot();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(0.7854, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void CschShouldReturnCorrectResult_Pi()
        {
            Csch? func = new Csch();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(Math.PI * -1);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-0.0866, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void RomanShouldReturnCorrectResult()
        {
            Roman? func = new Roman();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("II", result, "2 was not II");

            args = FunctionsHelper.CreateArgs(4);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("IV", result, "4 was not IV");

            args = FunctionsHelper.CreateArgs(14);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XIV", result, "14 was not XIV");

            args = FunctionsHelper.CreateArgs(23);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XXIII", result, "23 was not XXIII");

            args = FunctionsHelper.CreateArgs(59);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("LIX", result, "59 was not LIX");

            args = FunctionsHelper.CreateArgs(99);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XCIX", result, "99 was not XCIX");

            args = FunctionsHelper.CreateArgs(412);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("CDXII", result, "412 was not CDXII");

            args = FunctionsHelper.CreateArgs(1214);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MCCXIV", result, "1214 was not MCCXIV");

            args = FunctionsHelper.CreateArgs(3295);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MMMCCXCV", result, "3295 was not MMMCCXCV");
        }

        [TestMethod]
        public void RomanType1ShouldReturnCorrectResult()
        {
            Roman? func = new Roman();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(495, 1);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("LDVL", result, "495 was not LDVL");

            args = FunctionsHelper.CreateArgs(45, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VL", result, "45 was not VL");

            args = FunctionsHelper.CreateArgs(49, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VLIV", result, "59 was not VLIV");

            args = FunctionsHelper.CreateArgs(99, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VCIV", result, "99 was not VCIV");

            args = FunctionsHelper.CreateArgs(395, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("CCCVC", result, "395 was not CCCVC");

            args = FunctionsHelper.CreateArgs(949, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("CMVLIV", result, "949 was not CMVLIV");

            args = FunctionsHelper.CreateArgs(3295, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MMMCCVC", result, "3295 was not MMMCCVC");
        }

        [TestMethod]
        public void RomanType2ShouldReturnCorrectResult()
        {
            Roman? func = new Roman();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(495, 2);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XDV", result, "495 was not XDV");

            args = FunctionsHelper.CreateArgs(45, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VL", result, "45 was not VL");

            args = FunctionsHelper.CreateArgs(59, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("LIX", result, "59 was not LIX");

            args = FunctionsHelper.CreateArgs(99, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("IC", result, "99 was not IC");

            args = FunctionsHelper.CreateArgs(490, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XD", result, "490 was not XD");

            args = FunctionsHelper.CreateArgs(949, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("CMIL", result, "949 was not CMIL");

            args = FunctionsHelper.CreateArgs(2999, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MMXMIX", result, "2999 was not MMXMIX");
        }

        [TestMethod]
        public void RomanType3ShouldReturnCorrectResult()
        {
            Roman? func = new Roman();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(495, 3);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VD", result, "495 was not VD");

            args = FunctionsHelper.CreateArgs(499, 3);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VDIV", result, "499 was not VDIV");

            args = FunctionsHelper.CreateArgs(995, 3);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VM", result, "995 was not VM");

            args = FunctionsHelper.CreateArgs(999, 3);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VMIV", result, "999 was not VMIV");

            args = FunctionsHelper.CreateArgs(1999, 3);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MVMIV", result, "490 was not MVMIV");
        }

        [TestMethod]
        public void RomanType4ShouldReturnCorrectResult()
        {
            Roman? func = new Roman();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(499, 4);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("ID", result, "499 was not ID");

            args = FunctionsHelper.CreateArgs(999, 4);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("IM", result, "999 was not IM");
        }

        [TestMethod]
        public void Roman_should_work_correctly_with_all_possible_arguments_issue_770()
        {
            // This unit test was supplied via Github issue 
            using ExcelPackage? p = new ExcelPackage();
            p.Workbook.Worksheets.Add("first");

            ExcelWorksheet? sheet = p.Workbook.Worksheets.First();

            sheet.Cells["A1"].Formula = "=Roman(9.99)";
            sheet.Cells["A2"].Formula = "=Roman(999;1)";
            sheet.Cells["A3"].Formula = "=Roman(45;true)";
            sheet.Cells["A4"].Formula = "=Roman(499;false)";

            sheet.Calculate();

            // incorrect rounding method - have to be floor
            Assert.AreEqual("IX", sheet.Cells["A1"].Value.ToString());
            // incorrect cornversion by 1 scenario
            Assert.AreEqual("LMVLIV", sheet.Cells["A2"].Value.ToString());
            // incorrect interpretation of true - should be "0" instead of "1"
            Assert.AreEqual("XLV", sheet.Cells["A3"].Value.ToString());
            // incorrect interpretation of false - should be "4" instead of "0"
            Assert.AreEqual("ID", sheet.Cells["A4"].Value.ToString());
        }

        [TestMethod]
        public void GcdShouldReturnCorrectResult()
        {
            Gcd? func = new Gcd();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(15, 10, 25);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(5, result);

            args = FunctionsHelper.CreateArgs(0, 8, 12);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void LcmShouldReturnCorrectResult()
        {
            Lcm? func = new Lcm();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(15, 10, 25);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(150, result);

            args = FunctionsHelper.CreateArgs(1, 8, 12);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(24, result);
        }

        [TestMethod]
        public void SumShouldCalculate2Plus3AndReturn5()
        {
            Sum? func = new Sum();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, 3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumShouldCalculateEnumerableOf2Plus5Plus3AndReturn10()
        {
            Sum? func = new Sum();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void SumShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            Sum? func = new Sum();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3, 4);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void SumSqShouldCalculateArray()
        {
            Sumsq? func = new Sumsq();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, 4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(20d, result.Result);
        }

        [TestMethod]
        public void SumSqShouldIncludeTrueAsOne()
        {
            Sumsq? func = new Sumsq();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, 4, true);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(21d, result.Result);
        }

        [TestMethod]
        public void SumSqShouldNoCountTrueTrueInArray()
        {
            Sumsq? func = new Sumsq();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 4, true));
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(20d, result.Result);
        }

        [TestMethod]
        public void StdevShouldCalculateCorrectResult()
        {
            Stdev? func = new Stdev();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 3, 5);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void StdevaShouldCalculateCorrectResult()
        {
            Stdeva? func = new Stdeva();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 3, 5, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.7078d, System.Math.Round((double)result.Result, 4));

            args = FunctionsHelper.CreateArgs(1, 3, 5, 2, true, "text");
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.7889d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void StdevpaShouldCalculateCorrectResult()
        {
            Stdevpa? func = new Stdevpa();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 3, 5, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.479d, System.Math.Round((double)result.Result, 4));

            args = FunctionsHelper.CreateArgs(1, 3, 5, 2, true, "text");
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.633d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void StdevShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            Stdev? func = new Stdev();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 3, 5, 6);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void StdevPShouldCalculateCorrectResult()
        {
            StdevP? func = new StdevP();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, 3, 4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0.8165d, Math.Round((double)result.Result, 5));
        }

        [TestMethod]
        public void StdevPShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            StdevP? func = new StdevP();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, 3, 4, 165);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0.8165d, Math.Round((double)result.Result, 5));
        }

        [TestMethod]
        public void ExpShouldCalculateCorrectResult()
        {
            Exp? func = new Exp();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(54.59815003d, System.Math.Round((double)result.Result, 8));
        }

        [TestMethod]
        public void MaxShouldCalculateCorrectResult()
        {
            Max? func = new Max();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void MaxShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            Max? func = new Max();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            args.ElementAt(2).SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void MaxShouldHandleEmptyRange()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A5"].Formula = "MAX(A1:A4)";
            sheet.Calculate();
            object? value = sheet.Cells["A5"].Value;
            Assert.AreEqual(0d, value);
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResult()
        {
            Maxa? func = new Maxa();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-1, 0, 1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResultUsingBool()
        {
            Maxa? func = new Maxa();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-1, 0, true);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResultUsingString()
        {
            Maxa? func = new Maxa();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-1, "test");
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0d, result.Result);
        }

        [TestMethod]
        public void MinShouldCalculateCorrectResult()
        {
            Min? func = new Min();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void MinShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            Min? func = new Min();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, 2, 5, 3);
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void MinShouldHandleEmptyRange()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A5"].Formula = "MIN(A1:A4)";
            sheet.Calculate();
            object? value = sheet.Cells["A5"].Value;
            Assert.AreEqual(0d, value);
        }

        [TestMethod]
        public void AverageShouldCalculateCorrectResult()
        {
            double expectedResult = (4d + 2d + 5d + 2d) / 4d;
            Average? func = new Average();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldCalculateCorrectResultWithEnumerableAndBoolMembers()
        {
            double expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
            Average? func = new Average();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldIgnoreHiddenFieldsIfIgnoreHiddenValuesIsTrue()
        {
            double expectedResult = (4d + 2d + 2d + 1d) / 4d;
            Average? func = new Average();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldThrowDivByZeroExcelErrorValueIfEmptyArgs()
        {
            eErrorType errorType = eErrorType.Value;

            Average? func = new Average();
            FunctionArgument[]? args = new FunctionArgument[0];
            try
            {
                func.Execute(args, _parsingContext);
            }
            catch (ExcelErrorValueException e)
            {
                errorType = e.ErrorValue.Type;
            }
            Assert.AreEqual(eErrorType.Div0, errorType);
        }

        [TestMethod]
        public void AverageAShouldCalculateCorrectResult()
        {
            double expectedResult = (4d + 2d + 5d + 2d) / 4d;
            AverageA? func = new AverageA();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageAShouldIncludeTrueAs1()
        {
            double expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
            AverageA? func = new AverageA();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, true);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void AverageAShouldThrowValueExceptionIfNonNumericTextIsSupplied()
        {
            AverageA? func = new AverageA();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, "ABC");
            CompileResult? result = func.Execute(args, _parsingContext);
        }

        [TestMethod]
        public void AverageAShouldCountValueAs0IfNonNumericTextIsSuppliedInArray()
        {
            AverageA? func = new AverageA();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1d, 2d, 3d, "ABC"));
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.5d, result.Result);
        }

        [TestMethod]
        public void AverageAShouldCountNumericStringWithValue()
        {
            AverageA? func = new AverageA();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4d, 2d, "9");
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void RoundShouldReturnCorrectResult()
        {
            Round? func = new Round();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2.3433, 3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.343d, result.Result);
        }

        [TestMethod]
        public void RoundShouldReturnCorrectResultWhenNbrOfDecimalsIsNegative()
        {
            Round? func = new Round();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(9333, -3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9000d, result.Result);
        }

        [TestMethod]
        public void RandShouldReturnAValueBetween0and1()
        {
            Rand? func = new Rand();
            FunctionArgument[]? args = new FunctionArgument[0];
            CompileResult? result1 = func.Execute(args, _parsingContext);
            Assert.IsTrue(((double)result1.Result) > 0 && ((double)result1.Result) < 1);
            CompileResult? result2 = func.Execute(args, _parsingContext);
            Assert.AreNotEqual(result1.Result, result2.Result, "The two numbers were the same");
            Assert.IsTrue(((double)result2.Result) > 0 && ((double)result2.Result) < 1);
        }

        [TestMethod]
        public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValues()
        {
            RandBetween? func = new RandBetween();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 5);
            CompileResult? result = func.Execute(args, _parsingContext);
            CollectionAssert.Contains(new List<double> { 1d, 2d, 3d, 4d, 5d }, result.Result);
        }

        [TestMethod]
        public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValuesWhenLowIsNegative()
        {
            RandBetween? func = new RandBetween();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-5, 0);
            CompileResult? result = func.Execute(args, _parsingContext);
            CollectionAssert.Contains(new List<double> { 0d, -1d, -2d, -3d, -4d, -5d }, result.Result);
        }

        [TestMethod]
        public void CountShouldReturnNumberOfNumericItems()
        {
            Count? func = new Count();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4");
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void CountShouldIncludeNumericStringsAndDatesInArray()
        {
            Count? func = new Count();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4"));
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }


        [TestMethod]
        public void CountShouldIncludeEnumerableMembers()
        {
            Count? func = new Count();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod, Ignore]
        public void CountShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            Count? func = new Count();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountAShouldCountEmptyString()
        {
            CountA? func = new CountA();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4", null, string.Empty);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(6d, result.Result);
        }

        [TestMethod]
        public void CountAShouldIncludeEnumerableMembers()
        {
            CountA? func = new CountA();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void CountAShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            CountA? func = new CountA();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void ProductShouldMultiplyArguments()
        {
            Product? func = new Product();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2d, 2d, 4d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(16d, result.Result);
        }

        [TestMethod]
        public void ProductShouldHandleEnumerable()
        {
            Product? func = new Product();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void ProductShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            Product? func = new Product();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(16d, result.Result);
        }

        [TestMethod]
        public void ProductShouldHandleFirstItemIsEnumerable()
        {
            Product? func = new Product();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 2d, 2d);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void VarShouldReturnCorrectResult()
        {
            Var? func = new Var();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VaraShouldReturnCorrectResult()
        {
            Vara? func = new Vara();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 3, 5, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.9167d, System.Math.Round((double)result.Result, 4));

            args = FunctionsHelper.CreateArgs(1, 3, 5, 2, true, "text");
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3.2d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarpaShouldReturnCorrectResult()
        {
            Varpa? func = new Varpa();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 3, 5, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.1875, System.Math.Round((double)result.Result, 4));

            args = FunctionsHelper.CreateArgs(1, 3, 5, 2, true, "text");
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarDotSShouldReturnCorrectResult()
        {
            VarDotS? func = new VarDotS();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            Var? func = new Var();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 9);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarPShouldReturnCorrectResult()
        {
            VarP? func = new VarP();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.25d, result.Result);
        }

        [TestMethod]
        public void VarDotPShouldReturnCorrectResult()
        {
            VarDotP? func = new VarDotP();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.25d, result.Result);
        }

        [TestMethod]
        public void VarPShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            VarP? func = new VarP();
            func.IgnoreHiddenValues = true;
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 9);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.25d, result.Result);
        }

        [TestMethod]
        public void ModShouldReturnCorrectResult()
        {
            Mod? func = new Mod();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(5, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CosShouldReturnCorrectResult()
        {
            Cos? func = new Cos();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(-0.416146837d, roundedResult);
        }

        [TestMethod]
        public void CosHShouldReturnCorrectResult()
        {
            Cosh? func = new Cosh();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(3.762195691, roundedResult);
        }

        [TestMethod]
        public void AcosShouldReturnCorrectResult()
        {
            Acos? func = new Acos();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(0.1);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 4);
            Assert.AreEqual(1.4706, roundedResult);
        }

        [TestMethod]
        public void ACosHShouldReturnCorrectResult()
        {
            Acosh? func = new Acosh();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 3);
            Assert.AreEqual(1.317, roundedResult);
        }

        [TestMethod]
        public void SinShouldReturnCorrectResult()
        {
            Sin? func = new Sin();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.909297427, roundedResult);
        }

        [TestMethod]
        public void SinhShouldReturnCorrectResult()
        {
            Sinh? func = new Sinh();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(3.626860408d, roundedResult);
        }

        [TestMethod]
        public void TanShouldReturnCorrectResult()
        {
            Tan? func = new Tan();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(-2.185039863d, roundedResult);
        }

        [TestMethod]
        public void TanhShouldReturnCorrectResult()
        {
            Tanh? func = new Tanh();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.96402758d, roundedResult);
        }

        [TestMethod]
        public void AtanShouldReturnCorrectResult()
        {
            Atan? func = new Atan();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(10);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(1.471127674d, roundedResult);
        }

        [TestMethod]
        public void Atan2ShouldReturnCorrectResult()
        {
            Atan2? func = new Atan2();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(1.107148718d, roundedResult);
        }

        [TestMethod]
        public void AtanhShouldReturnCorrectResult()
        {
            Atanh? func = new Atanh();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(0.1);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 4);
            Assert.AreEqual(0.1003d, roundedResult);
        }

        [TestMethod]
        public void LogShouldReturnCorrectResult()
        {
            Log? func = new Log();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.301029996d, roundedResult);
        }

        [TestMethod]
        public void LogShouldReturnCorrectResultWithBase()
        {
            Log? func = new Log();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void Log10ShouldReturnCorrectResult()
        {
            Log10? func = new Log10();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.301029996d, roundedResult);
        }

        [TestMethod]
        public void LnShouldReturnCorrectResult()
        {
            Ln? func = new Ln();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(5);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 5);
            Assert.AreEqual(1.60944d, roundedResult);
        }

        [TestMethod]
        public void SqrtPiShouldReturnCorrectResult()
        {
            SqrtPi? func = new SqrtPi();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            CompileResult? result = func.Execute(args, _parsingContext);
            double roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(2.506628275d, roundedResult);
        }

        [TestMethod]
        public void SignShouldReturnMinus1IfArgIsNegative()
        {
            Sign? func = new Sign();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-2);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-1d, result);
        }

        [TestMethod]
        public void SignShouldReturn1IfArgIsPositive()
        {
            Sign? func = new Sign();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2);
            object? result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void RounddownShouldReturnCorrectResultWithPositiveNumber()
        {
            Rounddown? func = new Rounddown();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(9.999, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9.99, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleNegativeNumber()
        {
            Rounddown? func = new Rounddown();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-9.999, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-9.99, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleNegativeNumDigits()
        {
            Rounddown? func = new Rounddown();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(999.999, -2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(900d, result.Result);
        }

        [TestMethod]
        public void RounddownShouldReturn0IfNegativeNumDigitsIsTooLarge()
        {
            Rounddown? func = new Rounddown();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(999.999, -4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0d, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleZeroNumDigits()
        {
            Rounddown? func = new Rounddown();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(999.999, 0);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(999d, result.Result);
        }

        [TestMethod]
        public void RoundupShouldReturnCorrectResultWithPositiveNumber()
        {
            Roundup? func = new Roundup();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(9.9911, 3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9.992, result.Result);
        }

        [TestMethod]
        public void RoundupShouldHandleNegativeNumDigits()
        {
            Roundup? func = new Roundup();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(99123, -2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(99200d, result.Result);
        }

        [TestMethod]
        public void RoundupShouldHandleZeroNumDigits()
        {
            Roundup? func = new Roundup();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(999.999, 0);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1000d, result.Result);
        }

        [TestMethod]
        public void TruncShouldReturnCorrectResult()
        {
            Trunc? func = new Trunc();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(99.99);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(99d, result.Result);
        }

        [TestMethod]
        public void FactShouldRoundDownAndReturnCorrectResult()
        {
            Fact? func = new Fact();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(5.99);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(120d, result.Result);
        }

        [TestMethod]
        public void FactShouldReturnErrorNegativeNumber()
        {
            Fact? func = new Fact();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(-1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void FactDoubleShouldReturnCorrectResult5()
        {
            FactDouble? func = new FactDouble();
            IEnumerable<FunctionArgument>? arg = FunctionsHelper.CreateArgs(5);
            CompileResult? result = func.Execute(arg, _parsingContext);
            Assert.AreEqual(15d, result.Result);
        }

        [TestMethod]
        public void FactDoubleShouldReturnCorrectResult8()
        {
            FactDouble? func = new FactDouble();
            IEnumerable<FunctionArgument>? arg = FunctionsHelper.CreateArgs(8);
            CompileResult? result = func.Execute(arg, _parsingContext);
            Assert.AreEqual(384d, result.Result);
        }

        [TestMethod]
        public void QuotientShouldReturnCorrectResult()
        {
            Quotient? func = new Quotient();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(5, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void QuotientShouldReturnErrorDenomIs0()
        {
            Quotient? func = new Quotient();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 0);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void LargeShouldReturnTheLargestNumberIf1()
        {
            Large? func = new Large();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 3), 1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void LargeShouldReturnTheSecondLargestNumberIf2()
        {
            Large? func = new Large();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void LargeShouldReturnErrorIfIndexOutOfBounds()
        {
            Large? func = new Large();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void SmallShouldReturnTheSmallestNumberIf1()
        {
            Small? func = new Small();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 3), 1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SmallShouldReturnTheSecondSmallestNumberIf2()
        {
            Small? func = new Small();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void SmallShouldThrowIfIndexOutOfBounds()
        {
            Small? func = new Small();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void MedianShouldReturnErrorIfNoArgs()
        {
            Median? func = new Median();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.Empty();
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void MedianShouldCalculateCorrectlyWithOneMember()
        {
            Median? func = new Median();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void MedianShouldCalculateCorrectlyWithOddMembers()
        {
            Median? func = new Median();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(3, 5, 1, 4, 2);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void MedianShouldCalculateCorrectlyWithEvenMembers()
        {
            Median? func = new Median();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.5d, result.Result);
        }

        [TestMethod]
        public void CountIfShouldHandleNegativeCriteria()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = -1;
            sheet1.Cells["A2"].Value = -2;
            sheet1.Cells["A3"].Formula = "CountIf(A1:A2,\"-1\")";
            sheet1.Calculate();
            Assert.AreEqual(1d, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void OddShouldRound0To1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = 0;
            sheet1.Cells["A3"].Formula = "ODD(A1)";
            sheet1.Calculate();
            Assert.AreEqual(1, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void OddShouldRound1To1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = 1;
            sheet1.Cells["A3"].Formula = "ODD(A1)";
            sheet1.Calculate();
            Assert.AreEqual(1, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void OddShouldRound2To3()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = 2;
            sheet1.Cells["A3"].Formula = "ODD(A1)";
            sheet1.Calculate();
            Assert.AreEqual(3, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void OddShouldRoundMinus1point3ToMinus3()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = -1.3;
            sheet1.Cells["A3"].Formula = "ODD(A1)";
            sheet1.Calculate();
            Assert.AreEqual(-3, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void EvenShouldRound0To0()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = 0;
            sheet1.Cells["A3"].Formula = "EVEN(A1)";
            sheet1.Calculate();
            Assert.AreEqual(0, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void EvenShouldRound1To2()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = 1;
            sheet1.Cells["A3"].Formula = "EVEN(A1)";
            sheet1.Calculate();
            Assert.AreEqual(2, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void EvenShouldRound2To2()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = 2;
            sheet1.Cells["A3"].Formula = "EVEN(A1)";
            sheet1.Calculate();
            Assert.AreEqual(2, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void EvenShouldRoundMinus1point3ToMinus2()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("test");
            sheet1.Cells["A1"].Value = -1.3;
            sheet1.Cells["A3"].Formula = "EVEN(A1)";
            sheet1.Calculate();
            Assert.AreEqual(-2, sheet1.Cells["A3"].Value);
        }

        [TestMethod]
        public void Rank()
        {
            using ExcelPackage? p = new ExcelPackage();
            ExcelWorksheet? w = p.Workbook.Worksheets.Add("testsheet");
            w.SetValue(1, 1, 1);
            w.SetValue(2, 1, 1);
            w.SetValue(3, 1, 2);
            w.SetValue(4, 1, 2);
            w.SetValue(5, 1, 4);
            w.SetValue(6, 1, 4);

            w.SetFormula(1, 2, "RANK(1,A1:A5)");
            w.SetFormula(1, 3, "RANK(1,A1:A5,1)");
            w.SetFormula(1, 4, "RANK.AVG(1,A1:A5)");
            w.SetFormula(1, 5, "RANK.AVG(1,A1:A5,1)");

            w.SetFormula(2, 2, "RANK.EQ(2,A1:A5)");
            w.SetFormula(2, 3, "RANK.EQ(2,A1:A5,1)");
            w.SetFormula(2, 4, "RANK.AVG(2,A1:A5,1)");
            w.SetFormula(2, 5, "RANK.AVG(2,A1:A5,0)");

            w.SetFormula(3, 2, "RANK(3,A1:A5)");
            w.SetFormula(3, 3, "RANK(3,A1:A5,1)");
            w.SetFormula(3, 4, "RANK.AVG(3,A1:A5,1)");
            w.SetFormula(3, 5, "RANK.AVG(3,A1:A5,0)");

            w.SetFormula(4, 2, "RANK.EQ(4,A1:A5)");
            w.SetFormula(4, 3, "RANK.EQ(4,A1:A5,1)");
            w.SetFormula(4, 4, "RANK.AVG(4,A1:A5,1)");
            w.SetFormula(4, 5, "RANK.AVG(4,A1:A5)");


            w.SetFormula(5, 4, "RANK.AVG(4,A1:A6,1)");
            w.SetFormula(5, 5, "RANK.AVG(4,A1:A6)");

            w.Calculate();

            Assert.AreEqual(w.GetValue(1, 2), 4D);
            Assert.AreEqual(w.GetValue(1, 3), 1D);
            Assert.AreEqual(w.GetValue(1, 4), 4.5D);
            Assert.AreEqual(w.GetValue(1, 5), 1.5D);

            Assert.AreEqual(w.GetValue(2, 2), 2D);
            Assert.AreEqual(w.GetValue(2, 3), 3D);
            Assert.AreEqual(w.GetValue(2, 4), 3.5D);
            Assert.AreEqual(w.GetValue(2, 5), 2.5D);

            Assert.IsInstanceOfType(w.GetValue(3, 2), typeof(ExcelErrorValue));
            Assert.IsInstanceOfType(w.GetValue(3, 3), typeof(ExcelErrorValue));
            Assert.IsInstanceOfType(w.GetValue(3, 4), typeof(ExcelErrorValue));
            Assert.IsInstanceOfType(w.GetValue(3, 5), typeof(ExcelErrorValue));

            Assert.AreEqual(w.GetValue(4, 2), 1D);
            Assert.AreEqual(w.GetValue(4, 3), 5D);
            Assert.AreEqual(w.GetValue(4, 4), 5D);
            Assert.AreEqual(w.GetValue(4, 5), 1D);

            Assert.AreEqual(w.GetValue(5, 4), 5.5D);
            Assert.AreEqual(w.GetValue(5, 5), 1.5D);
        }

        [TestMethod]
        public void PercentrankInc_Test1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["A5"].Formula = "PERCENTRANK.INC(A1:A3,2)";
            sheet.Calculate();
            object? result = sheet.Cells["A5"].Value;
            Assert.AreEqual(0.5, result);

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 6;
            sheet.Cells["A5"].Formula = "PERCENTRANK.INC(A1:A3,3)";
            sheet.Calculate();
            result = sheet.Cells["A5"].Value;
            Assert.AreEqual(0.625, result);
        }

        [TestMethod]
        public void PercentrankInc_Test2()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 4;
            sheet.Cells["A4"].Value = 6.5;
            sheet.Cells["A5"].Value = 8;
            sheet.Cells["A6"].Value = 9;
            sheet.Cells["A7"].Value = 10;
            sheet.Cells["A8"].Value = 12;
            sheet.Cells["A9"].Value = 14;

            sheet.Cells["A10"].Formula = "PERCENTRANK.INC(A1:A9,6.5)";
            sheet.Calculate();
            object? result = sheet.Cells["A10"].Value;
            Assert.AreEqual(0.375, result);

            sheet.Cells["A10"].Formula = "PERCENTRANK.INC(A1:A9,7,5)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(0.41666, result);
        }

        [TestMethod] 
        public void PercentrankExc_Test1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 4;
            sheet.Cells["A4"].Value = 6.5;
            sheet.Cells["A5"].Value = 8;
            sheet.Cells["A6"].Value = 9;
            sheet.Cells["A7"].Value = 10;
            sheet.Cells["A8"].Value = 12;
            sheet.Cells["A9"].Value = 14;
            sheet.Cells["B1"].Formula = "PERCENTRANK.EXC(A1:A9,6.5)";
            sheet.Calculate();
            object? result = sheet.Cells["B1"].Value;
            Assert.AreEqual(0.4, result);

            sheet.Cells["B1"].Formula = "PERCENTRANK.EXC(A1:A9,7,5)";
            sheet.Calculate();
            result = sheet.Cells["B1"].Value;
            Assert.AreEqual(0.43333, result);

            sheet.Cells["B1"].Formula = "PERCENTRANK.EXC(A1:A9,18)";
            sheet.Calculate();
            result = sheet.Cells["B1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result);
        }

        [TestMethod]
        public void Percentile_Test1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 2;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 6;
            sheet.Cells["A4"].Value = 4;
            sheet.Cells["A5"].Value = 3;
            sheet.Cells["A6"].Value = 5;

            sheet.Cells["A10"].Formula = "PERCENTILE(A1:A6,0.2)";
            sheet.Calculate();
            object? result = sheet.Cells["A10"].Value;
            Assert.AreEqual(2d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE(A1:A6,60%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(4d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE(A1:A6,50%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(3.5d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE(A1:A6,95%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(5.75d, result);
        }

        [TestMethod]
        public void PercentileInc_Test1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 0;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 2;
            sheet.Cells["A4"].Value = 3;
            sheet.Cells["A5"].Value = 4;
            sheet.Cells["A6"].Value = 5;

            sheet.Cells["A10"].Formula = "PERCENTILE.INC(A1:A6,0.2)";
            sheet.Calculate();
            object? result = sheet.Cells["A10"].Value;
            Assert.AreEqual(1d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE.INC(A1:A6,60%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(3d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE.INC(A1:A6,50%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(2.5d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE.INC(A1:A6,95%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(4.75d, result);
        }

        [TestMethod]
        public void PercentileExc_Test1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["A4"].Value = 4;

            sheet.Cells["A10"].Formula = "PERCENTILE.EXC(A1:A4,0.2)";
            sheet.Calculate();
            object? result = sheet.Cells["A10"].Value;
            Assert.AreEqual(1d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE.EXC(A1:A4,60%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(3d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE.EXC(A1:A4,50%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(2.5d, result);

            sheet.Cells["A10"].Formula = "PERCENTILE.EXC(A1:A4,95%)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
        }

        [TestMethod]
        public void Quartile_Test1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 2;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 6;
            sheet.Cells["A4"].Value = 4;
            sheet.Cells["A5"].Value = 3;
            sheet.Cells["A6"].Value = 5;
            sheet.Cells["A7"].Value = 0;

            sheet.Cells["A10"].Formula = "QUARTILE(A1:A7,0)";
            sheet.Calculate();
            object? result = sheet.Cells["A10"].Value;
            Assert.AreEqual(0d, result);

            sheet.Cells["A10"].Formula = "QUARTILE(A1:A7,1)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(1.5d, result);

            sheet.Cells["A10"].Formula = "QUARTILE(A1:A7, 2)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(3d, result);

            sheet.Cells["A10"].Formula = "QUARTILE(A1:A7,3)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(4.5d, result);

            sheet.Cells["A10"].Formula = "QUARTILE(A1:A7,4)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(6d, result);
        }

        [TestMethod]
        public void QuartileInc_Test1()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 2;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 6;
            sheet.Cells["A4"].Value = 4;
            sheet.Cells["A5"].Value = 3;
            sheet.Cells["A6"].Value = 5;
            sheet.Cells["A7"].Value = 0;

            sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7,0)";
            sheet.Calculate();
            object? result = sheet.Cells["A10"].Value;
            Assert.AreEqual(0d, result);

            sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7,1)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(1.5d, result);

            sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7, 2)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(3d, result);

            sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7,3)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(4.5d, result);

            sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7,4)";
            sheet.Calculate();
            result = sheet.Cells["A10"].Value;
            Assert.AreEqual(6d, result);
        }

        [TestMethod]
        public void ModeShouldReturnCorrectResult()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 2;
            sheet.Cells["A4"].Value = 2;
            sheet.Cells["A5"].Value = 2;
            sheet.Cells["A6"].Value = 3;
            sheet.Cells["B1"].Formula = "MODE(A1:A6)";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void ModeShouldReturnLowestIfMultipleResults()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 2;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 2;
            sheet.Cells["A4"].Value = 1;
            sheet.Cells["A5"].Value = 3;
            sheet.Cells["A6"].Value = 3;
            sheet.Cells["B1"].Formula = "MODE.SNGL(A1:A6)";
            sheet.Calculate();
            Assert.AreEqual(1d, sheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void MultinomialShouldReturnCorrectResult()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 3;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 2;
            sheet.Cells["A4"].Value = 5;
            sheet.Cells["B1"].Formula = "MULTINOMIAL(A1:A4)";
            sheet.Calculate();
            Assert.AreEqual(27720d, sheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void CovarShouldReturnCorrectResult()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 3;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 2;
            sheet.Cells["A4"].Value = 5;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["B2"].Value = 6;
            sheet.Cells["B3"].Value = 2;
            sheet.Cells["B4"].Value = 8;

            sheet.Cells["C1"].Formula = "COVAR(A1:A4, B1:B4)";
            sheet.Calculate();
            Assert.AreEqual(1.625d, sheet.Cells["C1"].Value);

            sheet.Cells["C1"].Formula = "COVARIANCE.P(A1:A4, B1:B4)";
            sheet.Calculate();
            Assert.AreEqual(1.625d, sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void CovarianceSshouldReturnCorrectResult()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = 3;
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Value = 2;
            sheet.Cells["A4"].Value = 5;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["B2"].Value = 6;
            sheet.Cells["B3"].Value = 2;
            sheet.Cells["B4"].Value = 8;

            sheet.Cells["C1"].Formula = "COVARIANCE.S(A1:A4, B1:B4)";
            sheet.Calculate();
            Assert.AreEqual(2.16667d, System.Math.Round((double)sheet.Cells["C1"].Value, 5));
        }
    }
}
