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
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using System;
using System.IO;

namespace EPPlusTest.DataValidation;

[TestClass]
public class ListDataValidationTests : ValidationTestBase
{
    private IExcelDataValidationList _validation;

    [TestInitialize]
    public void Setup()
    {
        this.SetupTestData();
        this._validation = this._sheet.Workbook.Worksheets[1].DataValidations.AddListValidation("A1");
    }

    [TestCleanup]
    public void Cleanup()
    {
        this.CleanupTestData();
    }

    [TestMethod]
    public void ListDataValidation_FormulaIsSet()
    {
        Assert.IsNotNull(this._validation.Formula);
    }

    [TestMethod]
    public void ListDataValidation_CanAssignFormula()
    {
        this._validation.Formula.ExcelFormula = "abc!A2";
        Assert.AreEqual("abc!A2", this._validation.Formula.ExcelFormula);
    }

    [TestMethod]
    public void ListDataValidation_CanAssignDefinedName()
    {
        this._validation.Formula.ExcelFormula = "ListData";
        Assert.AreEqual("ListData", this._validation.Formula.ExcelFormula);
    }

    [TestMethod]
    public void ListDataValidation_WhenOneItemIsAddedCountIs1()
    {
        // Act
        this._validation.Formula.Values.Add("test");

        // Assert
        Assert.AreEqual(1, this._validation.Formula.Values.Count);
    }

    [TestMethod, ExpectedException(typeof(InvalidOperationException))]
    public void ListDataValidation_ShouldThrowWhenNoFormulaOrValueIsSet()
    {
        this._validation.Validate();
    }

    [TestMethod]
    public void ListDataValidation_ShowErrorMessageIsSet()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("list formula");

        ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].Value = "A";
        sheet2.Cells["A2"].Value = "B";
        sheet2.Cells["A3"].Value = "C";

        // add a validation and set values
        IExcelDataValidationList? validation = sheet.DataValidations.AddListValidation("A1");

        // Alternatively:
        // var validation = sheet.Cells["A1"].DataValidation.AddListDataValidation();
        validation.ShowErrorMessage = true;
        validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
        validation.ErrorTitle = "An invalid value was entered";
        validation.Error = "Select a value from the list";
        validation.Formula.ExcelFormula = "Sheet2!A1:A3";

        Assert.IsTrue(validation.ShowErrorMessage.Value);
    }

    [TestMethod]
    public void ListDataValidationExt_ShowDropDownIsSet()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("list formula");

        ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].Value = "A";
        sheet2.Cells["A2"].Value = "B";
        sheet2.Cells["A3"].Value = "C";

        // add a validation and set values
        IExcelDataValidationList? validation = sheet.DataValidations.AddListValidation("A1");

        // Alternatively:
        // var validation = sheet.Cells["A1"].DataValidation.AddListDataValidation();
        validation.HideDropDown = true;
        validation.ShowErrorMessage = true;
        validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
        validation.ErrorTitle = "An invalid value was entered";
        validation.Error = "Select a value from the list";
        validation.Formula.ExcelFormula = "Sheet2!A1:A3";

        // refresh the data validation
        validation = sheet.DataValidations.Find(x => x.Uid == validation.Uid).As.ListValidation;

        Assert.IsTrue(validation.HideDropDown.Value);
        ExcelDataValidationList? v = validation as ExcelDataValidationList;
        bool attributeValue = v.HideDropDown.Value;
        Assert.IsTrue(attributeValue);
    }

    [TestMethod]
    public void ListDataValidation_ShowDropDownIsSet()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("list formula");
        sheet.Cells["A1"].Value = "A";
        sheet.Cells["A2"].Value = "B";
        sheet.Cells["A3"].Value = "C";

        // add a validation and set values
        IExcelDataValidationList? validation = sheet.DataValidations.AddListValidation("B1");
        validation.HideDropDown = true;
        validation.ShowErrorMessage = true;
        validation.Formula.ExcelFormula = "A1:A3";

        Assert.IsTrue(validation.HideDropDown.Value);
        ExcelDataValidationList? v = validation as ExcelDataValidationList;
        bool attributeValue = v.HideDropDown.Value;
        Assert.IsTrue(attributeValue);
    }

    [TestMethod]
    public void ListDataValidation_AllowsOperatorShouldBeFalse()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("list operator");

        // add a validation and set values
        IExcelDataValidationList? validation = sheet.DataValidations.AddListValidation("A1");

        Assert.IsFalse(validation.AllowsOperator);
    }

    [TestMethod, ExpectedException(typeof(InvalidOperationException))]
    public void CompletelyEmptyListValidationsShouldThrow()
    {
        ExcelPackage? excel = new ExcelPackage();
        ExcelWorksheet? sheet = excel.Workbook.Worksheets.Add("Sheet1");
        sheet.Cells[1, 1].Value = "Column1";
        sheet.Cells["A2:A1048576"].DataValidation.AddListDataValidation();

        excel.Save();
    }

    [TestMethod]
    public void EmptyListValidationsShouldNotThrow()
    {
        ExcelPackage? excel = new ExcelPackage();
        ExcelWorksheet? sheet = excel.Workbook.Worksheets.Add("Sheet1");
        sheet.Cells[1, 1].Value = "Column1";
        IExcelDataValidationList? boolValidator = sheet.Cells["A2:A1048576"].DataValidation.AddListDataValidation();

        {
            boolValidator.Formula.Values.Add("");
            boolValidator.Formula.Values.Add("True");
            boolValidator.Formula.Values.Add("False");
            boolValidator.ShowErrorMessage = true;
            boolValidator.Error = "Choose either False or True";
        }

        excel.Save();
    }

    [TestMethod]
    public void ListLocalAndExt()
    {
        using ExcelPackage? package = OpenPackage("DataValidationExtLocalList.xlsx", true);
        ExcelWorksheet? ws1 = package.Workbook.Worksheets.Add("Worksheet1");
        package.Workbook.Worksheets.Add("Worksheet2");

        IExcelDataValidationDecimal? localVal = ws1.DataValidations.AddDecimalValidation("A1:A5");

        localVal.Formula.Value = 0;
        localVal.Formula2.Value = 0.1;

        IExcelDataValidationList? extVal = ws1.DataValidations.AddListValidation("B1:B5");

        extVal.Formula.ExcelFormula = "Worksheet2!$G$12:G15";

        SaveAndCleanup(package);

        ExcelPackage? p = OpenPackage("DataValidationExtLocalList.xlsx");

        MemoryStream? stream = new MemoryStream();
        p.SaveAs(stream);

        ExcelPackage pck = new ExcelPackage(stream);

        MemoryStream? stream2 = new MemoryStream();
        pck.SaveAs(stream2);
    }
}