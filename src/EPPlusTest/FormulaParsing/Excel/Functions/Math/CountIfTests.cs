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

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class CountIfTests
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
            this._parsingContext.Scopes.NewScope(RangeAddress.Empty);
            this._worksheet = this._package.Workbook.Worksheets.Add("testsheet");
        }

        [TestCleanup]
        public void Cleanup()
        {
            this._package.Dispose();
        }

        [TestMethod]
        public void CountIfNumeric()
        {
            this._worksheet.Cells["A1"].Value = 1d;
            this._worksheet.Cells["A2"].Value = 2d;
            this._worksheet.Cells["A3"].Value = 3d;
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, ">1");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfNonNumeric()
        {
            this._worksheet.Cells["A1"].Value = "Monday";
            this._worksheet.Cells["A2"].Value = "Tuesday";
            this._worksheet.Cells["A3"].Value = "Thursday";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "T*day");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        public void CountIfNullExpression()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = 1d;
            this._worksheet.Cells["A3"].Value = null;
            this._worksheet.Cells["B2"].Value = null;
            CountIf? func = new CountIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 2, 2, 2, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(0d, result.Result);
        }

        [TestMethod]
        public void CountIfNumericExpression()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = 1d;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, 1d);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

[TestMethod]
        public void CountIfEqualToEmptyString()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfNotEqualToNull()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "<>");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfEqualToZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = 0d;
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "0");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfNotEqualToZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = 0d;
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "<>0");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfGreaterThanZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = 1d;
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, ">0");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfGreaterThanOrEqualToZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = 1d;
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, ">=0");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfLessThanZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = -1d;
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "<0");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfLessThanOrEqualToZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = -1d;
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "<=0");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfLessThanCharacter()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "<a");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfLessThanOrEqualToCharacter()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, "<=a");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfGreaterThanCharacter()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, ">a");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfGreaterThanOrEqualToCharacter()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, ">=a");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CountIfIgnoreNumericString()
        {
            this._worksheet.Cells["A1"].Value = "1";
            this._worksheet.Cells["A2"].Value = 2;
            this._worksheet.Cells["A3"].Value = "3";
            CountIf? func = new CountIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, ">0");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }
    }
}
