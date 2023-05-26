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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class InformationFunctionsTests
    {
        private ParsingContext _context;

        [TestInitialize]
        public void Setup()
        {
            _context = ParsingContext.Create();
        }

        [TestMethod]
        public void IsBlankShouldReturnTrueIfFirstArgIsNull()
        {
            IsBlank? func = new IsBlank();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(new object[]{null});
            CompileResult? result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsBlankShouldReturnTrueIfFirstArgIsEmptyString()
        {
            IsBlank? func = new IsBlank();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(string.Empty);
            CompileResult? result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsNumberShouldReturnTrueWhenArgIsNumeric()
        {
            IsNumber? func = new IsNumber();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1d);
            CompileResult? result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsNumberShouldReturnfalseWhenArgIsNonNumeric()
        {
            IsNumber? func = new IsNumber();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs("1");
            CompileResult? result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsErrorShouldReturnTrueIfArgIsAnErrorCode()
        {
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(ExcelErrorValue.Parse("#DIV/0!"));
            IsError? func = new IsError();
            CompileResult? result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsErrorShouldReturnFalseIfArgIsNotAnError()
        {
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs("A", 1);
            IsError? func = new IsError();
            CompileResult? result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsTextShouldReturnTrueWhenFirstArgIsAString()
        {
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs("1");
            IsText? func = new IsText();
            CompileResult? result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsTextShouldReturnFalseWhenFirstArgIsNotAString()
        {
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1);
            IsText? func = new IsText();
            CompileResult? result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsNonTextShouldReturnFalseWhenFirstArgIsAString()
        {
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs("1");
            IsNonText? func = new IsNonText();
            CompileResult? result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsNonTextShouldReturnTrueWhenFirstArgIsNotAString()
        {
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1);
            IsNonText? func = new IsNonText();
            CompileResult? result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsOddShouldReturnCorrectResult()
        {
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(3.123);
            IsOdd? func = new IsOdd();
            CompileResult? result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsEvenShouldReturnCorrectResult()
        {
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4.123);
            IsEven? func = new IsEven();
            CompileResult? result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsLogicalShouldReturnCorrectResult()
        {
            IsLogical? func = new IsLogical();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1);
            CompileResult? result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);

            args = FunctionsHelper.CreateArgs("true");
            result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);

            args = FunctionsHelper.CreateArgs(false);
            result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void NshouldReturnCorrectResult()
        {
            N? func = new N();

            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1.2);
            CompileResult? result = func.Execute(args, _context);
            Assert.AreEqual(1.2, result.Result);

            args = FunctionsHelper.CreateArgs("abc");
            result = func.Execute(args, _context);
            Assert.AreEqual(0d, result.Result);

            args = FunctionsHelper.CreateArgs(true);
            result = func.Execute(args, _context);
            Assert.AreEqual(1d, result.Result);

            ExcelErrorValue? errorCode = ExcelErrorValue.Create(eErrorType.Value);
            args = FunctionsHelper.CreateArgs(errorCode);
            result = func.Execute(args, _context);
            Assert.AreEqual(errorCode, result.Result);
        }

        [TestMethod]
        public void TypeShouldReturn1WhenNumber()
        {
            using(ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "TYPE(2)";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void TypeShouldReturn1WhenEmpty()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "TYPE(A2)";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void TypeShouldReturn2WhenText()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "TYPE(A2)";
                sheet.Cells["A2"].Value = "asdf";
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void TypeShouldReturn4WhenBool()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "TYPE(A2)";
                sheet.Cells["A2"].Value = true;
                sheet.Calculate();
                Assert.AreEqual(4, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void TypeShouldReturn16WhenError()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "TYPE(A2)";
                sheet.Cells["A2"].Formula = "1/0";
                sheet.Calculate();
                Assert.AreEqual(16, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void TypeShouldReturn64WhenArray()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "TYPE({1,2,3})";
                sheet.Calculate();
                Assert.AreEqual(64, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SheetShouldReturnCorrectResult()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "SHEET()";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);

                package.Workbook.Worksheets.Add("Sheet2");
                sheet.Cells["A1"].Formula = "SHEET(\"Sheet2\")";
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SheetShouldReturnCorrectResult_Ref()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("Sheet2");

                sheet.Cells["A1"].Formula = "SHEET(Sheet2!A1)";
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SheetShouldReturnCorrectResult_SheetName()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("Sheet2");

                sheet.Names.Add("aName", sheet2.Cells["B1:C3"]);
                sheet.Cells["A1"].Formula = "SHEET(aName)";
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SheetShouldReturnCorrectResult_WbName()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("Sheet2");

                package.Workbook.Names.Add("aName", sheet2.Cells["B1:C3"]);
                sheet.Cells["A1"].Formula = "SHEET(aName)";
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SheetShouldReturnCorrectResult_Table()
        {
            // TODO: support table names as expressions in formula calc engine
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("Sheet2");

                sheet2.Tables.Add(sheet2.Cells["D1:G5"], "myTable");
                sheet.Cells["A1"].Formula = "SHEET(myTable)";
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["A1"].Value);
            }
        }
    }
}
