using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class MinAndMaxifsTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Init()
        {
            this._package = new ExcelPackage();
            ExcelWorksheet? sheet = this._package.Workbook.Worksheets.Add("test");
            sheet.Cells["B3"].Value = "Hannah";
            sheet.Cells["C3"].Value = "F";
            sheet.Cells["D3"].Value = 93;
            sheet.Cells["B4"].Value = "Edward";
            sheet.Cells["C4"].Value = "M";
            sheet.Cells["D4"].Value = 79;
            sheet.Cells["B5"].Value = "Miranda";
            sheet.Cells["C5"].Value = "F";
            sheet.Cells["D5"].Value = 85;
            sheet.Cells["B6"].Value = "Miranda";
            sheet.Cells["C6"].Value = "F";
            sheet.Cells["D6"].Value = 82;
            sheet.Cells["B7"].Value = "William";
            sheet.Cells["C7"].Value = "M";
            sheet.Cells["D7"].Value = 64;
            this._worksheet = sheet;
        }

        [TestCleanup]
        public void Cleanup()
        {
            this._package.Dispose();
        }

        [TestMethod]
        public void MaxIfsShouldHandleOneCriteria()
        {
            this._worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"F\")";
            this._worksheet.Calculate();
            Assert.AreEqual(93d, this._worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldHandleTwoCriterias()
        {
            this._worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"F\", B3:B7, \"Miranda\")";
            this._worksheet.Calculate();
            Assert.AreEqual(85d, this._worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldHandleTwoCriteriasWithWildcards()
        {
            this._worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"F\", B3:B7, \"Mi**nda\")";
            this._worksheet.Calculate();
            Assert.AreEqual(85d, this._worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldReturnZeroIfNoMatch()
        {
            this._worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"P\")";
            this._worksheet.Calculate();
            Assert.AreEqual(0d, this._worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MaxIfsShouldReturnValueErrorIfWrongSizeOnCriteriaRange()
        {
            this._worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7,C3:C7,\"F\", B3:B5, \"Mi**nda\")";
            this._worksheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value).ToString(), this._worksheet.Cells["F1"].Value.ToString());
        }

        [TestMethod]
        public void MaxIfsShouldHandleNumericCriteriaWithOperator()
        {
            this._worksheet.Cells["F1"].Formula = "MAXIFS(D3:D7, D3:D7,\">0\")";
            this._worksheet.Calculate();
            Assert.AreEqual(93d, this._worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldHandleOneCriteria()
        {
            this._worksheet.Cells["F1"].Formula = "MINIFS(D3:D7,C3:C7,\"F\")";
            this._worksheet.Calculate();
            Assert.AreEqual(82d, this._worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldHandleTwoCriterias()
        {
            this._worksheet.Cells["F1"].Formula = "MINIFS(D3:D7,C3:C7,\"F\", B3:B7, \"Miranda\")";
            this._worksheet.Calculate();
            Assert.AreEqual(82d, this._worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldHandleTwoCriteriasWithWildcards()
        {
            this._worksheet.Cells["F1"].Formula = "MINIFS(D3:D7,C3:C7,\"F\", B3:B7, \"Mi**nda\")";
            this._worksheet.Calculate();
            Assert.AreEqual(82d, this._worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MinIfsShouldReturnZeroIfNoMatch()
        {
            this._worksheet.Cells["F1"].Formula = "MINIFS(D3:D7,C3:C7,\"P\")";
            this._worksheet.Calculate();
            Assert.AreEqual(0d, this._worksheet.Cells["F1"].Value); ;
        }

        [TestMethod]
        public void MinIfsShouldHandleNumericCriteriaWithOperator()
        {
            this._worksheet.Cells["F1"].Formula = "MINIFS(D3:D7, D3:D7,\">0\")";
            this._worksheet.Calculate();
            Assert.AreEqual(64d, this._worksheet.Cells["F1"].Value);
        }
    }
}
