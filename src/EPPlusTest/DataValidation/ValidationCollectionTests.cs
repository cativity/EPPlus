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
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;

namespace EPPlusTest.DataValidation;

[TestClass]
public class ValidationCollectionTests : ValidationTestBase
{
    [TestInitialize]
    public void Setup() => this.SetupTestData();

    [TestCleanup]
    public void Cleanup() => this.CleanupTestData();

    [TestMethod, ExpectedException(typeof(ArgumentNullException))]
    public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenAddressIsNullOrEmpty() =>

        // Act
        _ = this._sheet.DataValidations.AddDecimalValidation(string.Empty);

    [TestMethod, ExpectedException(typeof(InvalidOperationException))]
    public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenNewValidationCollidesWithExisting()
    {
        // Act
        _ = this._sheet.DataValidations.AddDecimalValidation("A1");
        _ = this._sheet.DataValidations.AddDecimalValidation("A1");
    }

    [TestMethod, ExpectedException(typeof(InvalidOperationException))]
    public void ExcelDataValidationCollection_AddInteger_ShouldThrowWhenNewValidationCollidesWithExisting()
    {
        // Act
        _ = this._sheet.DataValidations.AddIntegerValidation("A1");
        _ = this._sheet.DataValidations.AddIntegerValidation("A1");
    }

    [TestMethod, ExpectedException(typeof(InvalidOperationException))]
    public void ExcelDataValidationCollection_AddTextLength_ShouldThrowWhenNewValidationCollidesWithExisting()
    {
        // Act
        _ = this._sheet.DataValidations.AddTextLengthValidation("A1");
        _ = this._sheet.DataValidations.AddTextLengthValidation("A1");
    }

    [TestMethod, ExpectedException(typeof(InvalidOperationException))]
    public void ExcelDataValidationCollection_AddDateTime_ShouldThrowWhenNewValidationCollidesWithExisting()
    {
        // Act
        _ = this._sheet.DataValidations.AddDateTimeValidation("A1");
        _ = this._sheet.DataValidations.AddDateTimeValidation("A1");
    }

    [TestMethod]
    public void ExcelDataValidationCollection_Index_ShouldReturnItemAtIndex()
    {
        // Arrange
        _ = this._sheet.DataValidations.AddDateTimeValidation("A1");
        _ = this._sheet.DataValidations.AddDateTimeValidation("A2");
        _ = this._sheet.DataValidations.AddDateTimeValidation("B1");

        // Act
        ExcelDataValidation? result = this._sheet.DataValidations[1];

        // Assert
        Assert.AreEqual("A2", result.Address.Address);
    }

    [TestMethod]
    public void ExcelDataValidationCollection_FindAll_ShouldReturnValidationInColumnAonly()
    {
        // Arrange
        _ = this._sheet.DataValidations.AddDateTimeValidation("A1");
        _ = this._sheet.DataValidations.AddDateTimeValidation("A2");
        _ = this._sheet.DataValidations.AddDateTimeValidation("B1");

        // Act
        IEnumerable<ExcelDataValidation>? result = this._sheet.DataValidations.FindAll(x => x.Address.Address.StartsWith("A"));

        // Assert
        Assert.AreEqual(2, result.Count());
    }

    [TestMethod]
    public void ExcelDataValidationCollection_Find_ShouldReturnFirstMatchOnly()
    {
        // Arrange
        _ = this._sheet.DataValidations.AddDateTimeValidation("A1");
        _ = this._sheet.DataValidations.AddDateTimeValidation("A2");

        // Act
        ExcelDataValidation? result = this._sheet.DataValidations.Find(x => x.Address.Address.StartsWith("A"));

        // Assert
        Assert.AreEqual("A1", result.Address.Address);
    }

    [TestMethod]
    public void ExcelDataValidationCollection_Clear_ShouldBeEmpty()
    {
        // Arrange
        _ = this._sheet.DataValidations.AddDateTimeValidation("A1");

        // Act
        this._sheet.DataValidations.Clear();

        // Assert
        Assert.AreEqual(0, this._sheet.DataValidations.Count);
    }

    [TestMethod]
    public void ExcelDataValidationCollection_ExtLst_Clear_ShouldBeEmpty()
    {
        // Arrange
        _ = this._package.Workbook.Worksheets.Add("Sheet2");
        IExcelDataValidationList? v = this._sheet.DataValidations.AddListValidation("A1");
        v.Formula.ExcelFormula = "Sheet2!A1:A2";

        // Act
        this._sheet.DataValidations.Clear();

        // Assert
        Assert.AreEqual(0, this._sheet.DataValidations.Count);
    }

    [TestMethod]
    public void ExcelDataValidationCollection_RemoveAll_ShouldRemoveMatchingEntries()
    {
        // Arrange
        _ = this._sheet.DataValidations.AddIntegerValidation("A1");
        _ = this._sheet.DataValidations.AddIntegerValidation("A2");
        _ = this._sheet.DataValidations.AddIntegerValidation("B1");

        // Act
        this._sheet.DataValidations.RemoveAll(x => x.Address.Address.StartsWith("B"));

        // Assert
        Assert.AreEqual(2, this._sheet.DataValidations.Count);
    }
}