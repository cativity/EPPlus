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
    public class AverageIfTests
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
        public void AverageIfNumeric()
        {
            this._worksheet.Cells["A1"].Value = 1d;
            this._worksheet.Cells["A2"].Value = 2d;
            this._worksheet.Cells["A3"].Value = 3d;
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">1", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void AverageIfNonNumeric()
        {
            this._worksheet.Cells["A1"].Value = "Monday";
            this._worksheet.Cells["A2"].Value = "Tuesday";
            this._worksheet.Cells["A3"].Value = "Thursday";
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "T*day", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void AverageIfNumericExpression()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = 1d;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            AverageIf? func = new AverageIf();
            IRangeInfo range = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range, 1d);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void AverageIfEqualToEmptyString()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void AverageIfNotEqualToNull()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<>", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void AverageIfEqualToZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = 0d;
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "0", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void AverageIfNotEqualToZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = 0d;
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<>0", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void AverageIfGreaterThanZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = 1d;
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">0", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void AverageIfGreaterThanOrEqualToZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = 1d;
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">=0", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void AverageIfLessThanZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = -1d;
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<0", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void AverageIfLessThanOrEqualToZero()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = -1d;
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<=0", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void AverageIfLessThanCharacter()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<a", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void AverageIfLessThanOrEqualToCharacter()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, "<=a", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void AverageIfGreaterThanCharacter()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">a", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void AverageIfGreaterThanOrEqualToCharacter()
        {
            this._worksheet.Cells["A1"].Value = null;
            this._worksheet.Cells["A2"].Value = string.Empty;
            this._worksheet.Cells["A3"].Value = "Not Empty";
            this._worksheet.Cells["B1"].Value = 1d;
            this._worksheet.Cells["B2"].Value = 3d;
            this._worksheet.Cells["B3"].Value = 5d;
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = this._provider.GetRange(this._worksheet.Name, 1, 2, 3, 2);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void AverageIfShouldNotCompareNumericStrings()
        {
            this._worksheet.Cells["A1"].Value = "1";
            this._worksheet.Cells["A2"].Value = 2;
            this._worksheet.Cells["A3"].Value = "3";
            AverageIf? func = new AverageIf();
            IRangeInfo range1 = this._provider.GetRange(this._worksheet.Name, 1, 1, 3, 1);
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(range1, ">0");
            CompileResult? result = func.Execute(args, this._parsingContext);
            Assert.AreEqual(2d, result.Result);
        }
    }
}
