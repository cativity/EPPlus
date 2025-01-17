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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Filter;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel.Functions;

[TestClass]
public class SubtotalTests
{
    private ParsingContext _context;

    [TestInitialize]
    public void Setup()
    {
        this._context = ParsingContext.Create();
        _ = this._context.Scopes.NewScope(RangeAddress.Empty);
    }

    [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
    public void ShouldThrowIfInvalidFuncNumber()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(139, 1);
        _ = func.Execute(args, this._context);
    }

    [TestMethod]
    public void ShouldCalculateAverageWhenCalcTypeIs1()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(1, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(30d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateCountWhenCalcTypeIs2()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateCountAWhenCalcTypeIs3()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(3, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateMaxWhenCalcTypeIs4()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(4, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(50d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateMinWhenCalcTypeIs5()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(5, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(10d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateProductWhenCalcTypeIs6()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(6, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(12000000d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateStdevWhenCalcTypeIs7()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(7, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        double resultRounded = Math.Round((double)result.Result, 5);
        Assert.AreEqual(15.81139d, resultRounded);
    }

    [TestMethod]
    public void ShouldCalculateStdevPWhenCalcTypeIs8()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(8, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        double resultRounded = Math.Round((double)result.Result, 8);
        Assert.AreEqual(14.14213562, resultRounded);
    }

    [TestMethod]
    public void ShouldCalculateSumWhenCalcTypeIs9()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(9, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(150d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateVarWhenCalcTypeIs10()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(10, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(250d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateVarPWhenCalcTypeIs11()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(11, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(200d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateAverageWhenCalcTypeIs101()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(101, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(30d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateCountWhenCalcTypeIs102()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(102, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateCountAWhenCalcTypeIs103()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(103, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(5d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateMaxWhenCalcTypeIs104()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(104, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(50d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateMinWhenCalcTypeIs105()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(105, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(10d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateProductWhenCalcTypeIs106()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(106, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(12000000d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateStdevWhenCalcTypeIs107()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(107, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        double resultRounded = Math.Round((double)result.Result, 5);
        Assert.AreEqual(15.81139d, resultRounded);
    }

    [TestMethod]
    public void ShouldCalculateStdevPWhenCalcTypeIs108()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(108, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        double resultRounded = Math.Round((double)result.Result, 8);
        Assert.AreEqual(14.14213562, resultRounded);
    }

    [TestMethod]
    public void ShouldCalculateSumWhenCalcTypeIs109()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(109, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(150d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateVarWhenCalcTypeIs110()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(110, 10, 20, 30, 40, 50, 51);
        args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(250d, result.Result);
    }

    [TestMethod]
    public void ShouldCalculateVarPWhenCalcTypeIs111()
    {
        Subtotal? func = new Subtotal();
        IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(111, 10, 20, 30, 40, 50);
        CompileResult? result = func.Execute(args, this._context);
        Assert.AreEqual(200d, result.Result);
    }

    [TestMethod]
    public void ShouldHandleMultipleLevelsOfSubtotals()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet3 = package.Workbook.Worksheets.Add("sheet3");
        sheet3.Cells["A1"].Value = 26959.64;
        sheet3.Cells["A2"].Value = 82272d;
        sheet3.Cells["A3"].Formula = "SUBTOTAL(9,A1:A2)";
        sheet3.Cells["A4"].Formula = "SUBTOTAL(9,A1:A3)";

        ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("sheet2");
        sheet2.Cells["A1"].Formula = "sheet3!A4";
        package.Workbook.Calculate();
        Assert.AreEqual(109231.64d, sheet2.Cells["A1"].Value);

        sheet3.Cells["A3"].Formula = "SUBTOTAL(8,A1:A2)";
        sheet3.Cells["A4"].Formula = "SUBTOTAL(8,A1:A3)";
        package.Workbook.Calculate();
        Assert.AreEqual(27656.18, sheet2.Cells["A1"].Value);

        sheet3.Cells["A3"].Formula = "SUBTOTAL(7,A1:A2)";
        sheet3.Cells["A4"].Formula = "SUBTOTAL(7,A1:A3)";
        package.Workbook.Calculate();
        Assert.AreEqual(39111.7448d, Math.Round((double)sheet2.Cells["A1"].Value, 4));
    }

    [TestMethod]
    public void ShouldHandleAutoFilters()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("sheet1");
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["A2"].Value = "data 1";
        sheet.Cells["A3"].Value = "data 2";
        sheet.Cells["A4"].Value = "data 3";
        sheet.Cells["A5"].Value = "data 4";
        sheet.Cells["A6"].Value = "data 5";

        sheet.Cells["B1"].Value = "Amount";
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["B3"].Value = 100;
        sheet.Cells["B4"].Value = 100;
        sheet.Cells["B5"].Value = 100;
        sheet.Cells["B6"].Value = 100;
        sheet.Cells["B7"].Formula = "SUBTOTAL(9,B2:B6)";
        sheet.Cells["A1:B6"].AutoFilter = true;
        ExcelValueFilterColumn? col = sheet.AutoFilter.Columns.AddValueFilterColumn(0);
        _ = col.Filters.Add(new ExcelFilterValueItem("data 1"));
        _ = col.Filters.Add(new ExcelFilterValueItem("data 3"));
        _ = col.Filters.Add(new ExcelFilterValueItem("data 4"));
        sheet.AutoFilter.ApplyFilter();

        Assert.IsFalse(sheet.Row(2).Hidden);
        Assert.IsTrue(sheet.Row(3).Hidden);
        Assert.IsFalse(sheet.Row(4).Hidden);
        Assert.IsFalse(sheet.Row(5).Hidden);
        Assert.IsTrue(sheet.Row(6).Hidden);

        package.Workbook.Calculate();
        Assert.AreEqual(300d, Math.Round((double)sheet.Cells["B7"].Value, 4));
    }
}