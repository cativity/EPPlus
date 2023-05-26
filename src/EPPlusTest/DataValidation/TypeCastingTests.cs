using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.DataValidation.Contracts;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class TypeCastingTests
    {
        [TestMethod]
        public void ShouldCastListValidation()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.DataValidations.AddListValidation("A1");
            IExcelDataValidationList? dv = sheet.DataValidations.First().As.ListValidation;
            Assert.IsNotNull(dv);
        }

        [TestMethod]
        public void ShouldCastIntegerValidation()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.DataValidations.AddIntegerValidation("A1");
            IExcelDataValidationInt? dv = sheet.DataValidations.First().As.IntegerValidation;
            Assert.IsNotNull(dv);
        }

        [TestMethod]
        public void ShouldCastDecimalValidation()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.DataValidations.AddDecimalValidation("A1");
            IExcelDataValidationDecimal? dv = sheet.DataValidations.First().As.DecimalValidation;
            Assert.IsNotNull(dv);
        }

        [TestMethod]
        public void ShouldCastDateTimeValidation()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.DataValidations.AddDateTimeValidation("A1");
            IExcelDataValidationDateTime? dv = sheet.DataValidations.First().As.DateTimeValidation;
            Assert.IsNotNull(dv);
        }

        [TestMethod]
        public void ShouldCastTimeValidation()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.DataValidations.AddTimeValidation("A1");
            IExcelDataValidationTime? dv = sheet.DataValidations.First().As.TimeValidation;
            Assert.IsNotNull(dv);
        }

        [TestMethod]
        public void ShouldCastCustomValidation()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.DataValidations.AddCustomValidation("A1");
            IExcelDataValidationCustom? dv = sheet.DataValidations.First().As.CustomValidation;
            Assert.IsNotNull(dv);
        }

        [TestMethod]
        public void ShouldCastAnyValidation()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.DataValidations.AddAnyValidation("A1");
            IExcelDataValidationAny? dv = sheet.DataValidations.First().As.AnyValidation;
            Assert.IsNotNull(dv);
        }
    }
}
