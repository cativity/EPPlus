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
using OfficeOpenXml.Utils;
using OfficeOpenXml;

namespace EPPlusTest.Utils;

[TestClass]
public class AddressUtilityTests
{
    [TestMethod]
    public void ParseForEntireColumnSelections_ShouldAddMaxRows()
    {
        // Arrange
        string? address = "A:A";

        // Act
        string? result = AddressUtility.ParseEntireColumnSelections(address);

        // Assert
        Assert.AreEqual("A1:A" + ExcelPackage.MaxRows, result);
    }

    [TestMethod]
    public void ParseForEntireColumnSelections_ShouldAddMaxRowsOnColumnsWithMultipleLetters()
    {
        // Arrange
        string? address = "AB:AC";

        // Act
        string? result = AddressUtility.ParseEntireColumnSelections(address);

        // Assert
        Assert.AreEqual("AB1:AC" + ExcelPackage.MaxRows, result);
    }

    [TestMethod]
    public void ParseForEntireColumnSelections_ShouldHandleMultipleRanges()
    {
        // Arrange
        string? address = "A:A B:B";
        string? expected = string.Format("A1:A{0} B1:B{0}", ExcelPackage.MaxRows);

        // Act
        string? result = AddressUtility.ParseEntireColumnSelections(address);

        // Assert
        Assert.AreEqual(expected, result);
    }

    [TestMethod]
    public void ShouldShiftRowInFormulaAddresses1()
    {
        string? formula = "SUM(A3:A4)";
        string? ws = "test";
        string? result = AddressUtility.ShiftAddressRowsInFormula(ws, formula, 3, -2);
        Assert.AreEqual("SUM(A1:A2)", result);
    }

    [TestMethod]
    public void ShouldNotShiftRowInFormulaFixedAddresses()
    {
        string? formula = "SUM(A$3:A$4)";
        string? ws = "test";
        string? result = AddressUtility.ShiftAddressRowsInFormula(ws, formula, 3, -2);
        Assert.AreEqual("SUM(A$3:A$4)", result);
    }
}