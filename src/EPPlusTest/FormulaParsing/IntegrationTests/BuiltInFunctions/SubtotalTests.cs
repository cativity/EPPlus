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
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions;

[TestClass]
public class SubtotalTests : FormulaParserTestBase
{
    private ExcelWorksheet _worksheet;
    private ExcelPackage _package;

    [TestInitialize]
    public void Setup()
    {
        this._package = new ExcelPackage(new MemoryStream());
        this._worksheet = this._package.Workbook.Worksheets.Add("Test");
        this._parser = this._worksheet.Workbook.FormulaParser;
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._package.Dispose();
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_Avg()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(1, A2:A3)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
        this._worksheet.Cells["A3"].Value = 2d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(2d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_Count()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(2, A2:A3)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
        this._worksheet.Cells["A3"].Value = 2d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(1d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_CountA()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(3, A2:A3)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
        this._worksheet.Cells["A3"].Value = 2d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(1d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_Max()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(4, A2:A3)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
        this._worksheet.Cells["A3"].Value = 2d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(2d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_Min()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(5, A2:A3)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
        this._worksheet.Cells["A3"].Value = 2d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(2d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_Product()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(6, A2:A3)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
        this._worksheet.Cells["A3"].Value = 2d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(2d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_Stdev()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(7, A2:A4)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(7, A5:A6)";
        this._worksheet.Cells["A3"].Value = 5d;
        this._worksheet.Cells["A4"].Value = 4d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Cells["A6"].Value = 4d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        result = Math.Round((double)result, 9);
        Assert.AreEqual(0.707106781d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_StdevP()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(8, A2:A4)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
        this._worksheet.Cells["A3"].Value = 5d;
        this._worksheet.Cells["A4"].Value = 4d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Cells["A6"].Value = 4d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(0.5d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_Sum()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A3)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
        this._worksheet.Cells["A3"].Value = 2d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(2d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_Var()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A4)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
        this._worksheet.Cells["A3"].Value = 5d;
        this._worksheet.Cells["A4"].Value = 4d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Cells["A6"].Value = 4d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(9d, result);
    }

    [TestMethod]
    public void SubtotalShouldNotIncludeSubtotalChildren_VarP()
    {
        this._worksheet.Cells["A1"].Formula = "SUBTOTAL(10, A2:A4)";
        this._worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
        this._worksheet.Cells["A3"].Value = 5d;
        this._worksheet.Cells["A4"].Value = 4d;
        this._worksheet.Cells["A5"].Value = 2d;
        this._worksheet.Cells["A6"].Value = 4d;
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A1"].Value;
        Assert.AreEqual(0.5d, result);
    }
}