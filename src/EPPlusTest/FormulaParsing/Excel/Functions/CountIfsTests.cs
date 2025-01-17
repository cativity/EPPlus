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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.Excel.Functions;

[TestClass]
public class CountIfsTests
{
    private ExcelPackage _package;
    private ExcelWorksheet _worksheet;

    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        this._worksheet = this._package.Workbook.Worksheets.Add("testsheet");
    }

    [TestCleanup]
    public void Cleanup() => this._package.Dispose();

    [TestMethod]
    public void ShouldHandleSingleNumericCriteria()
    {
        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Value = 1;
        this._worksheet.Cells["A3"].Value = 2;
        this._worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, 1)";
        this._worksheet.Calculate();
        Assert.AreEqual(2d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void ShouldHandleSingleRangeCriteria()
    {
        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Value = 1;
        this._worksheet.Cells["A3"].Value = 2;
        this._worksheet.Cells["B1"].Value = 1;
        this._worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, B1)";
        this._worksheet.Calculate();
        Assert.AreEqual(2d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void ShouldHandleSingleNumericWildcardCriteria()
    {
        this._worksheet.Cells["A1"].Value = 1;
        this._worksheet.Cells["A2"].Value = 2;
        this._worksheet.Cells["A3"].Value = 3;
        this._worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, \"<3\")";
        this._worksheet.Calculate();
        Assert.AreEqual(2d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void ShouldHandleSingleStringCriteria()
    {
        this._worksheet.Cells["A1"].Value = "abc";
        this._worksheet.Cells["A2"].Value = "def";
        this._worksheet.Cells["A3"].Value = "def";
        this._worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, \"def\")";
        this._worksheet.Calculate();
        Assert.AreEqual(2d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void ShouldHandleSingleStringWildcardCriteria()
    {
        this._worksheet.Cells["A1"].Value = "abc";
        this._worksheet.Cells["A2"].Value = "def";
        this._worksheet.Cells["A3"].Value = "def";
        this._worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, \"d*f\")";
        this._worksheet.Calculate();
        Assert.AreEqual(2d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void ShouldHandleSingleStringWildcardCriteriaStartingWildcard()
    {
        this._worksheet.Cells["A1"].Value = "abc";
        this._worksheet.Cells["A2"].Value = "def";
        this._worksheet.Cells["A3"].Value = "def";
        this._worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, \"*ef\")";
        this._worksheet.Calculate();
        Assert.AreEqual(2d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void ShouldHandleNullRangeCriteria()
    {
        this._worksheet.Cells["A1"].Value = null;
        this._worksheet.Cells["A2"].Value = 1;
        this._worksheet.Cells["A3"].Value = null;
        this._worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, B1)";
        this._worksheet.Calculate();
        Assert.AreEqual(0d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void ShouldIgnoreCellsWithErrors()
    {
        this._worksheet.Cells["A1"].Formula = "1/0";
        this._worksheet.Cells["A2"].Value = 1;
        this._worksheet.Cells["A3"].Value = null;
        this._worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, \">0\")";
        this._worksheet.Calculate();
        Assert.AreEqual(1d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void ShouldHandleMultipleRangesAndCriterias()
    {
        this._worksheet.Cells["A1"].Value = "abc";
        this._worksheet.Cells["A2"].Value = "def";
        this._worksheet.Cells["A3"].Value = "def";
        this._worksheet.Cells["A4"].Value = "def";
        this._worksheet.Cells["B1"].Value = 1;
        this._worksheet.Cells["B2"].Value = 2;
        this._worksheet.Cells["B3"].Value = 3;
        this._worksheet.Cells["B4"].Value = 2;
        this._worksheet.Cells["C1"].Value = null;
        this._worksheet.Cells["C2"].Value = 200;
        this._worksheet.Cells["C3"].Value = 3;
        this._worksheet.Cells["C4"].Value = 2;
        this._worksheet.Cells["A5"].Formula = "COUNTIFS(A1:A4, \"d*f\", B1:B4; 2; C1:C4; 200)";
        this._worksheet.Calculate();
        Assert.AreEqual(1d, this._worksheet.Cells["A5"].Value);
    }

    [TestMethod]
    public void CountIfs_CountThisRowWithoutCircularReferences()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
        sheet1.Cells["A1"].Value = "SumResult";

        // This shouldn't be a circular reference, because the 1:1="COUNTABLE" condition should filter out A2 before the 2:2 filter is applied
        sheet1.Cells["A2"].Formula = "COUNTIFS(1:1,\"COUNTABLE\",2:2,\"<>\")";

        sheet1.Cells["B2"].Value = 1;
        sheet1.Cells["C2"].Value = 2;
        sheet1.Cells["D2"].Value = 3;
        sheet1.Cells["E2"].Value = 4;
        sheet1.Cells["F2"].Value = 5;
        sheet1.Cells["G2"].Value = 6;

        sheet1.Cells["C1"].Value = "COUNTABLE";
        sheet1.Cells["D1"].Value = "COUNTABLE";
        sheet1.Cells["E1"].Value = "COUNTABLE";
        sheet1.Cells["G1"].Value = "COUNTABLE";

        pck.Workbook.Calculate(x => x.AllowCircularReferences = true);

        Assert.AreEqual(4, sheet1.Cells["A2"].GetValue<double>(), double.Epsilon);
    }

    [TestMethod]
    public void CountIfs_CountThisColWithoutCircularReferences()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
        sheet1.Cells["A1"].Value = "SumResult";

        // This shouldn't be a circular reference, because the 1:1="COUNTABLE" condition should filter out A2 before the 2:2 filter is applied
        sheet1.Cells["B1"].Formula = "COUNTIFS(A:A,\"COUNTABLE\",B:B,\"<>\")";

        sheet1.Cells["B2"].Value = 1;
        sheet1.Cells["B3"].Value = 2;
        sheet1.Cells["B4"].Value = 3;
        sheet1.Cells["B5"].Value = 4;
        sheet1.Cells["B6"].Value = 5;
        sheet1.Cells["B7"].Value = 6;

        sheet1.Cells["A3"].Value = "COUNTABLE";
        sheet1.Cells["A4"].Value = "COUNTABLE";
        sheet1.Cells["A5"].Value = "COUNTABLE";
        sheet1.Cells["A7"].Value = "COUNTABLE";

        pck.Workbook.Calculate(x => x.AllowCircularReferences = true);

        Assert.AreEqual(4, sheet1.Cells["B1"].GetValue<double>(), double.Epsilon);
    }
}