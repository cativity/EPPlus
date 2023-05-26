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
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges;

[TestClass]
public class TextExcelRangeTests
{
    private ExcelPackage _package;
    private ExcelWorksheet _worksheet;
    private CultureInfo _currentCulture;

    [TestInitialize]
    public void Initialize()
    {
        this._currentCulture = CultureInfo.CurrentCulture;
        this._package = new ExcelPackage();
        this._worksheet = this._package.Workbook.Worksheets.Add("Test");

        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Value = 3;
        this._worksheet.Cells["A3"].Value = 6;
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._package.Dispose();
        Thread.CurrentThread.CurrentCulture = this._currentCulture;
    }

    [TestMethod]
    public void ExactShouldReturnTrueWhenEqualValues()
    {
        this._worksheet.Cells["A2"].Value = 1d;
        this._worksheet.Cells["A4"].Formula = "EXACT(A1,A2)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.IsTrue((bool)result);
    }

    [TestMethod]
    public void FindShouldReturnIndexCaseSensitive()
    {
        this._worksheet.Cells["A1"].Value = "h";
        this._worksheet.Cells["A2"].Value = "Hej hopp";
        this._worksheet.Cells["A4"].Formula = "Find(A1,A2)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(5, result);
    }

    [TestMethod]
    public void FindShouldUse1basedIndex()
    {
        this._worksheet.Cells["A4"].Formula = "Find(\"P\",\"P2\",1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(1, result);
    }

    [TestMethod]
    public void SearchShouldReturnIndexCaseInSensitive()
    {
        this._worksheet.Cells["A1"].Value = "h";
        this._worksheet.Cells["A2"].Value = "Hej hopp";
        this._worksheet.Cells["A4"].Formula = "Search(A1,A2)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(1, result);
    }

    [TestMethod]
    public void SearchShouldUse1basedIndex()
    {
        this._worksheet.Cells["A4"].Formula = "Search(\"P\",\"P2\",1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(1, result);
    }

    [TestMethod]
    public void ValueShouldHandleStringWithIntegers()
    {
        this._worksheet.Cells["A1"].Value = "12";
        this._worksheet.Cells["A4"].Formula = "Value(A1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(12d, result);
    }

    [TestMethod]
    public void ValueShouldHandle1000delimiter()
    {
        string? delimiter = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
        string? val = $"5{delimiter}000";
        this._worksheet.Cells["A1"].Value = val;
        this._worksheet.Cells["A4"].Formula = "Value(A1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(5000d, result);
    }

    [TestMethod]
    public void ValueShouldHandle1000DelimiterAndDecimal()
    {
        string? delimiter = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
        string? decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        string? val = $"5{delimiter}000{decimalSeparator}123";
        this._worksheet.Cells["A1"].Value = val;
        this._worksheet.Cells["A4"].Formula = "Value(A1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(5000.123d, result);
    }

    [TestMethod]
    public void ValueShouldHandlePercent()
    {
        string? val = $"20%";
        this._worksheet.Cells["A1"].Value = val;
        this._worksheet.Cells["A4"].Formula = "Value(A1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(0.2d, result);
    }

    [TestMethod]
    public void ValueShouldHandleScientificNotation()
    {
        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
        this._worksheet.Cells["A1"].Value = "1.2345E-02";
        this._worksheet.Cells["A4"].Formula = "Value(A1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(0.012345d, result);
    }

    [TestMethod]
    public void ValueShouldHandleDate()
    {
        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
        DateTime date = new DateTime(2015, 12, 31);
        this._worksheet.Cells["A1"].Value = date.ToString(CultureInfo.CurrentCulture);
        this._worksheet.Cells["A4"].Formula = "Value(A1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(date.ToOADate(), result);
    }

    [TestMethod]
    public void ValueShouldHandleTime()
    {
        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
        DateTime date = new DateTime(2015, 12, 31);
        DateTime date2 = new DateTime(2015, 12, 31, 12, 00, 00);
        TimeSpan ts = date2.Subtract(date);
        this._worksheet.Cells["A1"].Value = ts.ToString();
        this._worksheet.Cells["A4"].Formula = "Value(A1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(0.5, result);
    }

    [TestMethod]
    public void ValueShouldReturn0IfValueIsNull()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A4"].Formula = "Value(A1)";
        this._worksheet.Calculate();
        object? result = this._worksheet.Cells["A4"].Value;
        Assert.AreEqual(0d, result);
    }

}