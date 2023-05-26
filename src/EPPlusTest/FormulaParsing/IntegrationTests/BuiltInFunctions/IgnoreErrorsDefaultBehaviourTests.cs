using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class IgnoreErrorsDefaultBehaviourTests
    {
        [TestMethod]
        public void SumShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "SUM(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void AverageShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "AVERAGE(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void AverageAShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "AVERAGEA(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void MaxShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "MAX(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void MinShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "MIN(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void MedianShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "MEDIAN(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void LargeShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "LARGE(B1:B2,ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void SmallShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "SMALL(B1:B2,ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void StdevDotSshouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "STDEV.S(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void StdevDotPshouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "STDEV.P(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void ProductShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "PRODUCT(ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void PercentileIncShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "PERCENTILE.INC(B1:B2,ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }

        [TestMethod]
        public void PercentileExcShouldReturnNameIfItContainsUnknownFunction()
        {
            using ExcelPackage? package = new ExcelPackage();
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Formula = "PERCENTILE.EXC(B1:B2,ABC(1))";
            sheet.Calculate();
            object? val = sheet.Cells["A1"].Value;
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), val);
        }
    }
}
