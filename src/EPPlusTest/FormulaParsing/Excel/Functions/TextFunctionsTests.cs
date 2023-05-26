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
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading;

namespace EPPlusTest.Excel.Functions.Text
{
    [TestClass]
    public class TextFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [TestMethod]
        public void CStrShouldConvertNumberToString()
        {
            CStr? func = new CStr();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(1), _parsingContext);
            Assert.AreEqual(DataType.String, result.DataType);
            Assert.AreEqual("1", result.Result);
        }

        [TestMethod]
        public void LenShouldReturnStringsLength()
        {
            Len? func = new Len();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("abc"), _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void LowerShouldReturnLowerCaseString()
        {
            Lower? func = new Lower();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("ABC"), _parsingContext);
            Assert.AreEqual("abc", result.Result);
        }

        [TestMethod]
        public void UpperShouldReturnUpperCaseString()
        {
            Upper? func = new Upper();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("abc"), _parsingContext);
            Assert.AreEqual("ABC", result.Result);
        }

        [TestMethod]
        public void LeftShouldReturnSubstringFromLeft()
        {
            Left? func = new Left();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), _parsingContext);
            Assert.AreEqual("ab", result.Result);
        }

        [TestMethod]
        public void RightShouldReturnSubstringFromRight()
        {
            Right? func = new Right();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), _parsingContext);
            Assert.AreEqual("cd", result.Result);
        }

        [TestMethod]
        public void MidShouldReturnSubstringAccordingToParams()
        {
            Mid? func = new Mid();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("abcd", 1, 2), _parsingContext);
            Assert.AreEqual("ab", result.Result);
        }

        [TestMethod]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs1()
        {
            Replace? func = new Replace();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("testar", 1, 2, "hej"), _parsingContext);
            Assert.AreEqual("hejstar", result.Result);
        }

        [TestMethod]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs3()
        {
            Replace? func = new Replace();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("testar", 3, 3, "hej"), _parsingContext);
            Assert.AreEqual("tehejr", result.Result);
        }

        [TestMethod]
        public void SubstituteShouldReturnAReplacedStringAccordingToParamsWhen()
        {
            Substitute? func = new Substitute();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("testar testar", "es", "xx"), _parsingContext);
            Assert.AreEqual("txxtar txxtar", result.Result);
        }

        [TestMethod]
        public void ConcatenateShouldConcatenateThreeStrings()
        {
            Concatenate? func = new Concatenate();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("One", "Two", "Three"), _parsingContext);
            Assert.AreEqual("OneTwoThree", result.Result);
        }

        [TestMethod]
        public void ConcatenateShouldConcatenateStringWithInt()
        {
            Concatenate? func = new Concatenate();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(1, "Two"), _parsingContext);
            Assert.AreEqual("1Two", result.Result);
        }

        [TestMethod]
        public void ConcatShouldReturnValErrorIfMoreThan254Args()
        {
            Concat? func = new Concat();
            List<object> args = new List<object>();
            for (int i = 0; i < 255; i++)
            {
                args.Add("arg " + i);
            }
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(args.ToArray()), _parsingContext);
            Assert.AreEqual("#VALUE!", result.Result.ToString());
        }

        [TestMethod]
        public void ExactShouldReturnTrueWhenTwoEqualStrings()
        {
            Exact? func = new Exact();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("abc", "abc"), _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void ExactShouldReturnTrueWhenEqualStringAndDouble()
        {
            Exact? func = new Exact();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("1", 1d), _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void ExactShouldReturnFalseWhenStringAndNull()
        {
            Exact? func = new Exact();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("1", null), _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void ExactShouldReturnFalseWhenTwoEqualStringsWithDifferentCase()
        {
            Exact? func = new Exact();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("abc", "Abc"), _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void FindShouldReturnIndexOfFoundPhrase()
        {
            Find? func = new Find();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hej hopp"), _parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void FindShouldReturnIndexOfFoundPhraseBasedOnStartIndex()
        {
            Find? func = new Find();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hopp hopp", 2), _parsingContext);
            Assert.AreEqual(6, result.Result);
        }

        [TestMethod]
        public void ProperShouldSetFirstLetterToUpperCase()
        {
            Proper? func = new Proper();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("this IS A tEst.wi3th SOME w0rds östEr"), _parsingContext);
            Assert.AreEqual("This Is A Test.Wi3Th Some W0Rds Öster", result.Result);
        }

        [TestMethod]
        public void HyperLinkShouldReturnArgIfOneArgIsSupplied()
        {
            Hyperlink? func = new Hyperlink();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com"), _parsingContext);
            Assert.AreEqual("http://epplus.codeplex.com", result.Result);
        }

        [TestMethod]
        public void HyperLinkShouldReturnLastArgIfTwoArgsAreSupplied()
        {
            Hyperlink? func = new Hyperlink();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com", "EPPlus"), _parsingContext);
            Assert.AreEqual("EPPlus", result.Result);
        }

        [TestMethod]
        public void TrimShouldReturnDataTypeString()
        {
            Trim? func = new Trim();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(" epplus "), _parsingContext);
            Assert.AreEqual(DataType.String, result.DataType);
        }

        [TestMethod]
        public void TrimShouldTrimFromBothEnds()
        {
            Trim? func = new Trim();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(" epplus "), _parsingContext);
            Assert.AreEqual("epplus", result.Result);
        }

        [TestMethod]
        public void TrimShouldTrimMultipleSpaces()
        {
            Trim? func = new Trim();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(" epplus    5 "), _parsingContext);
            Assert.AreEqual("epplus 5", result.Result);
        }

        [TestMethod]
        public void CleanShouldReturnDataTypeString()
        {
            Clean? func = new Clean();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("epplus"), _parsingContext);
            Assert.AreEqual(DataType.String, result.DataType);
        }

        [TestMethod]
        public void CleanShouldRemoveNonPrintableChars()
        {
            StringBuilder? input = new StringBuilder();
            for (int x = 1; x < 32; x++)
            {
                input.Append((char)x);
            }
            input.Append("epplus");
            Clean? func = new Clean();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(input), _parsingContext);
            Assert.AreEqual("epplus", result.Result);
        }

        [TestMethod]
        public void UnicodeShouldReturnCorrectCode()
        {
            Unicode? func = new Unicode();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("B"), _parsingContext);
            Assert.AreEqual(66, result.Result);

            result = func.Execute(FunctionsHelper.CreateArgs("a"), _parsingContext);
            Assert.AreEqual(97, result.Result);
        }

        [TestMethod]
        public void UnicharShouldReturnCorrectChar()
        {
            Unichar? func = new Unichar();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(66), _parsingContext);
            Assert.AreEqual("B", result.Result);

            result = func.Execute(FunctionsHelper.CreateArgs(97), _parsingContext);
            Assert.AreEqual("a", result.Result);
        }

        [TestMethod]
        public void NumberValueShouldCastIntegerValue()
        {
            NumberValue? func = new NumberValue();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("1000"), _parsingContext);
            Assert.AreEqual(1000d, result.Result);
        }

        [TestMethod]
        public void NumberValueShouldCastDecinalValueWithCurrentCulture()
        {
            string? input = $"1{CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator}000{CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator}15";
            NumberValue? func = new NumberValue();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(input), _parsingContext);
            Assert.AreEqual(1000.15d, result.Result);
        }

        [TestMethod]
        public void NumberValueShouldCastDecinalValueWithSeparators()
        {
            string? input = $"1,000.15";
            NumberValue? func = new NumberValue();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(input, ".", ","), _parsingContext);
            Assert.AreEqual(1000.15d, result.Result);
        }

        [TestMethod]
        public void NumberValueShouldHandlePercentage()
        {
            string? input = $"1,000.15%";
            NumberValue? func = new NumberValue();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(input, ".", ","), _parsingContext);
            Assert.AreEqual(10.0015d, result.Result);
        }

        [TestMethod]
        public void NumberValueShouldHandleMultiplePercentage()
        {
            string? input = $"1,000.15%%";
            NumberValue? func = new NumberValue();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(input, ".", ","), _parsingContext);
            double r = System.Math.Round((double)result.Result, 15);
            Assert.AreEqual(0.100015d, r);
        }

        [TestMethod]
        public void TextjoinShouldReturnCorrectResult_IgnoreEmpty()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Value = "Hello";
            sheet.Cells["A2"].Value = "world";
            sheet.Cells["A3"].Value = "";
            sheet.Cells["A4"].Value = "!";
            sheet.Cells["A5"].Formula = "TEXTJOIN(\" \", TRUE, A1:A4)";
            sheet.Calculate();
            Assert.AreEqual("Hello world !", sheet.Cells["A5"].Value);
        }

        [TestMethod]
        public void TextjoinShouldReturnCorrectResult_AllowEmpty()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Value = "Hello";
            sheet.Cells["A2"].Value = "world";
            sheet.Cells["A3"].Value = "";
            sheet.Cells["A4"].Value = "!";
            sheet.Cells["A5"].Formula = "TEXTJOIN(\".\", False, A1:A4, \"how are you?\")";
            sheet.Calculate();
            Assert.AreEqual("Hello.world..!.how are you?", sheet.Cells["A5"].Value);
        }

        [TestMethod]
        public void DollarShouldReturnCorrectResult()
        {
            string? expected = 123.46.ToString("C2", CultureInfo.CurrentCulture);
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Value = 123.456;
            sheet.Cells["A2"].Formula = "DOLLAR(A1)";
            sheet.Calculate();
            Assert.AreEqual(expected, sheet.Cells["A2"].Value);

            expected = 123.5.ToString("C1", CultureInfo.CurrentCulture);
            sheet.Cells["A2"].Formula = "DOLLAR(A1, 1)";
            sheet.Calculate();
            Assert.AreEqual(expected, sheet.Cells["A2"].Value);

            expected = 123.ToString("C0", CultureInfo.CurrentCulture);
            sheet.Cells["A2"].Formula = "DOLLAR(A1, 0)";
            sheet.Calculate();
            Assert.AreEqual(expected, sheet.Cells["A2"].Value);

            expected = 120.ToString("C0", CultureInfo.CurrentCulture);
            sheet.Cells["A2"].Formula = "DOLLAR(A1, -1)";
            sheet.Calculate();
            Assert.AreEqual(expected, sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ValueShouldReturnCorrectResult()
        {
            CultureInfo? cc = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("EN-US");
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "1,234,567.89";
                sheet.Cells["A2"].Formula = "VALUE(A1)";
                sheet.Calculate();
                Assert.AreEqual(1234567.89, sheet.Cells["A2"].Value);
            }
            Thread.CurrentThread.CurrentCulture = cc;
        }
    }
}
