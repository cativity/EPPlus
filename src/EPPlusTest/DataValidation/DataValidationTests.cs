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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class DataValidationTests : ValidationTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            SetupTestData();
        }

        [TestCleanup]
        public void Cleanup()
        {
            CleanupTestData();
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsEqualAndFormula1IsEmpty()
        {
            IExcelDataValidationInt? validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.Operator = ExcelDataValidationOperator.equal;
            validation.Validate();
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadAllValidOperatorsOnAllTypes()
        {

        }

        public void TestTypeOperator(ExcelDataValidation type)
        {

        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void TestRangeAddMultipleTryAddingAfterShouldThrow()
        {
            ExcelPackage pck = OpenTemplatePackage("ValidationRangeTestMany.xlsx");

            ExcelDataValidationCollection? validations = pck.Workbook.Worksheets[0].DataValidations;

            validations.AddIntegerValidation("C8");
        }

        [TestMethod]
        public void TestRangeAddMultipleTryAddingAfterShouldNotThrow()
        {
            ExcelPackage pck = OpenTemplatePackage("ValidationRangeTestMany.xlsx");

            ExcelDataValidationCollection? validations = pck.Workbook.Worksheets[0].DataValidations;

            validations.AddIntegerValidation("Z8");
        }


        [TestMethod]
        public void TestRangeAddsMultipleInbetweenInstances()
        {
            ExcelPackage pck = OpenTemplatePackage("ValidationRangeTestMany.xlsx");

            ExcelDataValidationCollection? validations = pck.Workbook.Worksheets[0].DataValidations;

            StringBuilder sb = new StringBuilder();

            //Ensure all addresses exist in _validationsRD
            for (int i = 0; i < validations.Count; i++)
            {
                if (validations[i].Address.Addresses != null)
                {
                    List<ExcelAddressBase>? addresses = validations[i].Address.Addresses;

                    for (int j = 0; j < validations[i].Address.Addresses.Count; j++)
                    {
                        if (!validations._validationsRD.Exists(addresses[j]._fromRow, addresses[j]._fromCol, addresses[j]._toRow, addresses[j]._toCol))
                        {
                            sb.Append(addresses[j]+",");
                        }
                    }
                }
                else
                {
                    if (!validations._validationsRD.Exists(validations[i].Address._fromRow, validations[i].Address._fromCol, validations[i].Address._toRow, validations[i].Address._toCol))
                    {
                        sb.Append(validations[i].Address+",");
                    }
                }
            }

            Assert.AreEqual("", sb.ToString());
        }

        [TestMethod]
        public void TestRangeAddsSingularInstance()
        {
            ExcelPackage pck = OpenTemplatePackage("ValidationRangeTest.xlsx"); ;

            //pck.Workbook.Worksheets.Add("RangeTest");

            ExcelDataValidationCollection? validations = pck.Workbook.Worksheets[0].DataValidations;

            StringBuilder sb = new StringBuilder();

            //Ensure all addresses exist in _validationsRD
            for(int i = 0; i< validations.Count; i++) 
            {
                if(validations[i].Address.Addresses != null)
                {
                    List<ExcelAddressBase>? addresses = validations[i].Address.Addresses;

                    for (int j = 0; j < validations[i].Address.Addresses.Count; j++)
                    {
                        if(!validations._validationsRD.Exists(addresses[j]._fromRow, addresses[j]._fromCol, addresses[j]._toRow, addresses[j]._toCol))
                        {
                            sb.Append(addresses[i]);
                        }
                    }
                }
                else
                {
                    if (!validations._validationsRD.Exists(validations[i].Address._fromRow, validations[i].Address._fromCol, validations[i].Address._toRow, validations[i].Address._toCol))
                    {
                        sb.Append(validations[i].Address);
                    }
                }
            }

            Assert.AreEqual("",sb.ToString());
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadTypes()
        {
            ExcelPackage? P = new ExcelPackage(new MemoryStream());
            ExcelWorksheet? sheet = P.Workbook.Worksheets.Add("NewSheet");

            sheet.DataValidations.AddAnyValidation("A1");
            IExcelDataValidationInt? intDV = sheet.DataValidations.AddIntegerValidation("A2");
            intDV.Formula.Value = 1;
            intDV.Formula2.Value = 1;

            IExcelDataValidationDecimal? decimalDV = sheet.DataValidations.AddDecimalValidation("A3");

            decimalDV.Formula.Value = 1;
            decimalDV.Formula2.Value = 1;

            IExcelDataValidationList? listDV = sheet.DataValidations.AddListValidation("A4");

            listDV.Formula.Values.Add("5");
            listDV.Formula.Values.Add("Option");


            IExcelDataValidationInt? textDV = sheet.DataValidations.AddTextLengthValidation("A5");

            textDV.Formula.Value = 1;
            textDV.Formula2.Value = 1;

            IExcelDataValidationDateTime? dateTimeDV = sheet.DataValidations.AddDateTimeValidation("A6");

            dateTimeDV.Formula.Value = DateTime.MaxValue;
            dateTimeDV.Formula2.Value = DateTime.MinValue;

            IExcelDataValidationTime? timeDV = sheet.DataValidations.AddTimeValidation("A7");

            timeDV.Formula.Value.Hour = 1;
            timeDV.Formula2.Value.Hour = 2;

            IExcelDataValidationCustom? customValidation = sheet.DataValidations.AddCustomValidation("A8");
            customValidation.Formula.ExcelFormula = "A1+A2";

            MemoryStream xmlStream = new MemoryStream();
            P.SaveAs(xmlStream);

            ExcelPackage? P2 = new ExcelPackage(xmlStream);
            ExcelDataValidationCollection dataValidations = P2.Workbook.Worksheets[0].DataValidations;

            Assert.AreEqual(dataValidations[0].ValidationType.Type, eDataValidationType.Any);
            Assert.AreEqual(dataValidations[1].ValidationType.Type, eDataValidationType.Whole);
            Assert.AreEqual(dataValidations[2].ValidationType.Type, eDataValidationType.Decimal);
            Assert.AreEqual(dataValidations[3].ValidationType.Type, eDataValidationType.List);
            Assert.AreEqual(dataValidations[4].ValidationType.Type, eDataValidationType.TextLength);
            Assert.AreEqual(dataValidations[5].ValidationType.Type, eDataValidationType.DateTime);
            Assert.AreEqual(dataValidations[6].ValidationType.Type, eDataValidationType.Time);
            Assert.AreEqual(dataValidations[7].ValidationType.Type, eDataValidationType.Custom);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadOperator()
        {
            ExcelPackage? P = new ExcelPackage(new MemoryStream());
            ExcelWorksheet? sheet = P.Workbook.Worksheets.Add("NewSheet");

            IExcelDataValidationInt? validation = sheet.DataValidations.AddIntegerValidation("A1");
            validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 1;

            MemoryStream xmlStream = new MemoryStream();
            P.SaveAs(xmlStream);

            ExcelPackage? P2 = new ExcelPackage(xmlStream);
            Assert.AreEqual(P2.Workbook.Worksheets[0].DataValidations[0].Operator, ExcelDataValidationOperator.greaterThanOrEqual);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadShowErrorMessage()
        {
            ExcelPackage? P = new ExcelPackage(new MemoryStream());
            ExcelWorksheet? sheet = P.Workbook.Worksheets.Add("NewSheet");

            IExcelDataValidationInt? validation = sheet.DataValidations.AddIntegerValidation("A1");

            validation.ShowErrorMessage = true;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 1;

            MemoryStream xmlStream = new MemoryStream();
            P.SaveAs(xmlStream);

            ExcelPackage? P2 = new ExcelPackage(xmlStream);
            Assert.IsTrue(P2.Workbook.Worksheets[0].DataValidations[0].ShowErrorMessage);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadShowinputMessage()
        {
            ExcelPackage? package = new ExcelPackage(new MemoryStream());
            CreateSheetWithIntegerValidation(package).ShowInputMessage = true;
            Assert.IsTrue(ReadIntValidation(package).ShowInputMessage);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadPrompt()
        {
            ExcelPackage? package = new ExcelPackage(new MemoryStream());
            CreateSheetWithIntegerValidation(package).Prompt = "Prompt";
            Assert.AreEqual("Prompt", ReadIntValidation(package).Prompt);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadError()
        {
            ExcelPackage? package = new ExcelPackage(new MemoryStream());
            string? validation = CreateSheetWithIntegerValidation(package).Error = "Error";

            Assert.AreEqual("Error", ReadIntValidation(package).Error);
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadErrorTitle()
        {
            ExcelPackage? package = new ExcelPackage(new MemoryStream());
            CreateSheetWithIntegerValidation(package).ErrorTitle = "ErrorTitle";
            Assert.AreEqual("ErrorTitle", ReadIntValidation(package).ErrorTitle);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsBetweenAndFormula2IsEmpty()
        {
            IExcelDataValidationInt? validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.Formula.Value = 1;
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Validate();
        }

        [TestMethod]
        public void DataValidations_ShouldAcceptOneItemOnly()
        {
            IExcelDataValidationList? validation = _sheet.DataValidations.AddListValidation("A1");
            validation.Formula.Values.Add("1");
            validation.Validate();
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfAllowBlankIsNotSet()
        {
            IExcelDataValidationInt? validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.Validate();
        }

        [TestMethod]
        public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsOneColumn()
        {
            // Act
            IExcelDataValidationInt? validation = _sheet.DataValidations.AddIntegerValidation("A:A");

            // Assert
            Assert.AreEqual("A1:A" + ExcelPackage.MaxRows.ToString(), validation.Address.Address);
        }

        [TestMethod]
        public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsDifferentColumns()
        {
            // Act
            IExcelDataValidationInt? validation = _sheet.DataValidations.AddIntegerValidation("A:B");

            // Assert
            Assert.AreEqual(string.Format("A1:B{0}", ExcelPackage.MaxRows), validation.Address.Address);
        }
        [TestMethod]
        public void TestInsertRowsIntoVeryLongRangeWithDataValidation()
        {
            using (ExcelPackage? pck = new ExcelPackage())
            {
                // Add a sheet with data validation on the whole of column A except row 1
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
                string? dvAddress = "A2:A1048576";
                IExcelDataValidationCustom? dv = wks.DataValidations.AddCustomValidation(dvAddress);

                // Check that the data validation address was set correctly
                Assert.AreEqual(dvAddress, dv.Address.Address);

                // Insert some rows into the worksheet
                wks.InsertRow(5, 3);

                // Check that the data validation rule still applies to the same range (since there's nowhere to extend it to)
                Assert.AreEqual(dvAddress, dv.Address.Address);
            }
        }
        [TestMethod]
        public void TestInsertRowsAboveVeryLongRangeWithDataValidation()
        {
            using (ExcelPackage? pck = new ExcelPackage())
            {
                // Add a sheet with data validation on the whole of column A except rows 1-10
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
                string? dvAddress = "A11:A1048576";
                IExcelDataValidationAny? dv = wks.DataValidations.AddAnyValidation(dvAddress);

                // Check that the data validation address was set correctly
                Assert.AreEqual(dvAddress, dv.Address.Address);

                // Insert 3 rows into the worksheet above the data validation
                wks.InsertRow(5, 3);

                // Check that the data validation starts lower down, but ends in the same place
                Assert.AreEqual("A14:A1048576", dv.Address.Address);
            }
        }

        [TestMethod]
        public void TestInsertRowsToPushDataValidationOffSheet()
        {
            using (ExcelPackage? pck = new ExcelPackage())
            {
                // Add a sheet with data validation on the last two rows of column A
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
                string? dvAddress = "A1048575:A1048576";
                IExcelDataValidationCustom? dv = wks.DataValidations.AddCustomValidation(dvAddress);

                // Check that the data validation address was set correctly
                Assert.AreEqual(1, wks.DataValidations.Count);
                Assert.AreEqual(dvAddress, dv.Address.Address);

                // Insert enough rows into the worksheet above the data validation rule to push it off the sheet 
                wks.InsertRow(5, 10);

                // Check that the data validation rule no longer exists
                Assert.AreEqual(0, wks.DataValidations.Count);
            }
        }

        [TestMethod]
        public void TestLoadingWorksheet()
        {
            using (ExcelPackage? p = OpenTemplatePackage("DataValidationReadTest.xlsx"))
            {
                ExcelWorksheet? ws = p.Workbook.Worksheets[0];
                Assert.AreEqual(4, ws.DataValidations.Count);
            }
        }

        [TestMethod]
        public void DataValidationAny_AllowsOperatorShouldBeFalse()
        {
            using (ExcelPackage? pck = new ExcelPackage())
            {
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
                string? dvAddress = "A1";
                IExcelDataValidationAny? dv = wks.DataValidations.AddAnyValidation(dvAddress);

                Assert.IsFalse(dv.AllowsOperator);
            }
        }

        [TestMethod]
        public void DataValidationDefaults_AllowsOperatorShouldBeTrueOnCorrectTypes()
        {
            using (ExcelPackage? pck = new ExcelPackage())
            {
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");

                IExcelDataValidationInt? intValidation = wks.DataValidations.AddIntegerValidation("A1");
                IExcelDataValidationDecimal? decimalValidation = wks.DataValidations.AddDecimalValidation("A2");
                IExcelDataValidationInt? textLengthValidation = wks.DataValidations.AddTextLengthValidation("A3");
                IExcelDataValidationDateTime? dateTimeValidation = wks.DataValidations.AddDateTimeValidation("A4");
                IExcelDataValidationTime? timeValidation = wks.DataValidations.AddTimeValidation("A5");
                IExcelDataValidationCustom? customValidation = wks.DataValidations.AddCustomValidation("A6");

                Assert.IsTrue(intValidation.AllowsOperator);
                Assert.IsTrue(decimalValidation.AllowsOperator);
                Assert.IsTrue(textLengthValidation.AllowsOperator);
                Assert.IsTrue(dateTimeValidation.AllowsOperator);
                Assert.IsTrue(timeValidation.AllowsOperator);
                Assert.IsTrue(customValidation.AllowsOperator);
            }
        }

        [TestMethod]
        public void DataValidations_CloneShouldDeepCopy()
        {
            using (ExcelPackage? pck = new ExcelPackage())
            {
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");
                IExcelDataValidationInt? validation = wks.DataValidations.AddIntegerValidation("A1");
                ExcelDataValidation? clone = ((ExcelDataValidationInt)validation).GetClone();
                clone.Address = new ExcelDatavalidationAddress("A2", clone);

                Assert.AreNotEqual(validation.Address, clone.Address);
            }
        }

        [TestMethod]
        public void DataValidations_ShouldCopyAllProperties()
        {
            using (ExcelPackage? pck = new ExcelPackage())
            {
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");

                List<ExcelDataValidation> validations = new List<ExcelDataValidation>
                {
                    (ExcelDataValidation)wks.DataValidations.AddIntegerValidation("A1"),
                    (ExcelDataValidation)wks.DataValidations.AddDecimalValidation("A2"),
                    (ExcelDataValidation)wks.DataValidations.AddTextLengthValidation("A3"),
                    (ExcelDataValidation)wks.DataValidations.AddDateTimeValidation("A4"),
                    (ExcelDataValidation)wks.DataValidations.AddTimeValidation("A5"),
                    (ExcelDataValidation)wks.DataValidations.AddCustomValidation("A6"),
                    (ExcelDataValidation)wks.DataValidations.AddAnyValidation("A7"),
                    (ExcelDataValidation)wks.DataValidations.AddListValidation("A9")
                };

                foreach (ExcelDataValidation? validation in validations)
                {
                    validation.AllowBlank = true;
                    validation.Prompt = "prompt";
                    validation.PromptTitle = "promptTitle";
                    validation.Error = "error";
                    validation.ErrorTitle = "errorTitle";
                    validation.ShowInputMessage = true;
                    validation.ShowErrorMessage = true;
                    validation.ErrorStyle = ExcelDataValidationWarningStyle.information;

                    ExcelDataValidation? clone = validation.GetClone();

                    Assert.AreEqual(validation.AllowBlank, clone.AllowBlank);
                    Assert.AreEqual(validation.Prompt, clone.Prompt);
                    Assert.AreEqual(validation.Error, clone.Error);
                    Assert.AreEqual(validation.ErrorTitle, clone.ErrorTitle);
                    Assert.AreEqual(validation.ShowInputMessage, clone.ShowInputMessage);
                    Assert.AreEqual(validation.ShowErrorMessage, clone.ShowErrorMessage);
                    Assert.AreEqual(validation.ErrorStyle, clone.ErrorStyle);
                }
            }
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadIMEmode()
        {
            using (ExcelPackage? pck = OpenPackage("ImeTest.xlsx", true))
            {
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");

                IExcelDataValidationCustom? validation = wks.DataValidations.AddCustomValidation("A1");
                validation.ShowErrorMessage = true;
                validation.Formula.ExcelFormula = "=ISTEXT(A1)";
                validation.ImeMode = ExcelDataValidationImeMode.FullKatakana;

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void DataValidations_ShouldWriteReadIMEmodeAndWriteAgain()
        {
            using (ExcelPackage? pck = OpenPackage("ImeTestOFF.xlsx", true))
            {
                ExcelWorksheet? wks = pck.Workbook.Worksheets.Add("Sheet1");

                IExcelDataValidationCustom? validation = wks.DataValidations.AddCustomValidation("A1");
                validation.ShowErrorMessage = true;
                validation.Formula.ExcelFormula = "=ISTEXT(A1)";
                validation.ImeMode = ExcelDataValidationImeMode.Off;

                MemoryStream? stream = new MemoryStream();
                pck.SaveAs(stream);

                ExcelPackage? pck2 = new ExcelPackage(stream);

                pck2.SaveAs(stream);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void DataValidations_Insert_Test()
        {
            using (ExcelPackage? pck = OpenPackage("InsertTest.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("InsertTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A3");

                IExcelDataValidationDecimal? rangeValidation2 = ws.DataValidations.AddDecimalValidation("A52");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;

                rangeValidation.Address.Address = "A1,A3";

                IExcelDataValidationList? list = ws.DataValidations.AddListValidation("A2");

                list.Formula.Values.Add("TestValue");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidation()
        {
            using (ExcelPackage? pck = OpenPackage("ClearDataValidationTest.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A3");
                IExcelDataValidationInt? rangeValidation2 = ws.DataValidations.AddIntegerValidation("A4:A6");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;


                ws.Cells["A2"].DataValidation.ClearDataValidation();

                IExcelDataValidationList? list = ws.DataValidations.AddListValidation("A2");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidationAndAddressChangeWithSpacedAddresses()
        {
            using (ExcelPackage? pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A3 B5 C3 E15:E17");
                IExcelDataValidationInt? rangeValidation2 = ws.DataValidations.AddIntegerValidation("A4:A6");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;


                ws.Cells["A2:A3"].DataValidation.ClearDataValidation();
                ws.Cells["E16 A5"].DataValidation.ClearDataValidation();

                IExcelDataValidationList? list = ws.DataValidations.AddListValidation("A2");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidationAndAddressChangeWithSpacedAddressesViaCells()
        {
            using (ExcelPackage? pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");

                IExcelDataValidationList? listValidation = ws.Cells["A1:A3 B5 C3 E15:E17"].DataValidation.AddListDataValidation();

                listValidation.Formula.Values.Add("Value1");

                IExcelDataValidationInt? rangeValidation = ws.Cells["A4:A6"].DataValidation.AddIntegerDataValidation();

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;


                ws.Cells["A2:A3"].DataValidation.ClearDataValidation();
                ws.Cells["E16 A5"].DataValidation.ClearDataValidation();

                IExcelDataValidationList? list = ws.DataValidations.AddListValidation("A2");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidationOverARangeWithMultipleValidations()
        {
            using (ExcelPackage? pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A5");
                IExcelDataValidationInt? rangeValidation2 = ws.DataValidations.AddIntegerValidation("A6:A8");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;


                ws.Cells["A4:A7"].DataValidation.ClearDataValidation();

                IExcelDataValidationList? list = ws.DataValidations.AddListValidation("A4");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                Assert.AreEqual(rangeValidation.Address.Address, "A1:A3");
                Assert.AreEqual(rangeValidation2.Address.Address, "A8");

                SaveAndCleanup(pck);
            }
        }


        [TestMethod]
        public void ClearValidationOverARangeWithMultipleValidations2()
        {
            using (ExcelPackage? pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A5");
                IExcelDataValidationInt? rangeValidation2 = ws.DataValidations.AddIntegerValidation("A6:A8");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;

                ws.Cells["A3:A7"].DataValidation.ClearDataValidation();

                IExcelDataValidationList? list = ws.DataValidations.AddListValidation("A4");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");

                Assert.AreEqual(rangeValidation.Address.Address, "A1:A2");
                Assert.AreEqual(rangeValidation2.Address.Address, "A8");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ClearValidationOverBlockRanges()
        {
            using (ExcelPackage? pck = OpenPackage("ClearBlockRanges.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A1:D5");
                IExcelDataValidationInt? rangeValidation2 = ws.DataValidations.AddIntegerValidation("C6:C8");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;

                ws.Cells["A3:B7"].DataValidation.ClearDataValidation();

                IExcelDataValidationList? list = ws.DataValidations.AddListValidation("A4");
                IExcelDataValidationList? list2 = ws.DataValidations.AddListValidation("B3:B7");

                list.Formula.Values.Add("Value1");
                list.Formula.Values.Add("Value2");


                list2.Formula.Values.Add("Value21");
                list2.Formula.Values.Add("Value22");

                //Assert.AreEqual(rangeValidation.Address.Address, "A1:A2");
                //Assert.AreEqual(rangeValidation2.Address.Address, "A8");

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void DeleteRangeOneAddressTest()
        {
            using (ExcelPackage? pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A1:A5");
                IExcelDataValidationInt? rangeValidation2 = ws.DataValidations.AddIntegerValidation("A6:A8");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                rangeValidation2.Operator = ExcelDataValidationOperator.equal;
                rangeValidation2.Formula.Value = 6;

                ws.DeleteRow(2, 5);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ClearSingular()
        {
            using (ExcelPackage? pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A9");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                ws.Cells["A9"].DataValidation.ClearDataValidation();
            }
        }

        [TestMethod]
        public void ClearSingularSpaceSeparated()
        {
            using (ExcelPackage? pck = OpenPackage("ClearDataValidationTestAdress.xlsx", true))
            {
                ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("ClearTest");
                IExcelDataValidationInt? rangeValidation = ws.DataValidations.AddIntegerValidation("A9 A6 B12 C50");

                rangeValidation.Operator = ExcelDataValidationOperator.equal;
                rangeValidation.Formula.Value = 5;

                ws.Cells["A9"].DataValidation.ClearDataValidation();
                ws.Cells["B12"].DataValidation.ClearDataValidation();

                Assert.AreEqual("A6 C50", rangeValidation.Address.Address);
            }
        }

        [TestMethod]
        public void RemovalOfCellsAfterBeingRemovedAndAdded()
        {
            using (ExcelPackage? pck = OpenPackage("DataValidationsUserClearTest.xlsx", true))
            {
                ExcelWorksheet? myWS = pck.Workbook.Worksheets.Add("MyWorksheet");
                ExcelWorksheet? yourWS = pck.Workbook.Worksheets.Add("YourWorksheet");

                IExcelDataValidationInt? validation = myWS.DataValidations.AddTextLengthValidation("A1:C5");

                validation.Operator = ExcelDataValidationOperator.lessThan;

                validation.Formula.Value = 10;

                myWS.Cells["B3:C6"].DataValidation.ClearDataValidation();

                IExcelDataValidationDecimal? decimalVal = myWS.Cells["B3:D4"].DataValidation.AddDecimalDataValidation();
                decimalVal.Operator = ExcelDataValidationOperator.greaterThan;
                decimalVal.Formula.Value = 5;

                myWS.Cells["B1:D2"].DataValidation.ClearDataValidation();

                SaveAndCleanup(pck);
            }
        }


        [TestMethod]
        public void UserTestClear()
        {
            using (ExcelPackage? pck = OpenPackage("DataValidationsUserClearTest.xlsx", true))
            {
                ExcelWorksheet? myWS = pck.Workbook.Worksheets.Add("MyWorksheet");
                ExcelWorksheet? yourWS = pck.Workbook.Worksheets.Add("YourWorksheet");

                IExcelDataValidationInt? validation = myWS.DataValidations.AddTextLengthValidation("A1:E30");

                validation.Operator = ExcelDataValidationOperator.lessThan;

                validation.Formula.Value = 10;

                myWS.Cells["C1:D10"].DataValidation.ClearDataValidation();

                IExcelDataValidationList? listVal = myWS.Cells["C5:D10"].DataValidation.AddListDataValidation();

                listVal.Formula.ExcelFormula = "$C$1:$D$4";

                myWS.Cells["B1:C4"].DataValidation.ClearDataValidation();

                SaveAndCleanup(pck);
            }
        }

        //C11:D30
    }
}
