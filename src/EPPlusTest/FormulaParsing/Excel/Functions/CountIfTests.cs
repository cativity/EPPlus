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
    public class CountIfTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            this._package = new ExcelPackage();
            this._sheet = this._package.Workbook.Worksheets.Add("Sheet1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            this._package.Dispose();
        }

        [TestMethod]
        public void CountIfShouldCalculateOneCriteria()
        {
            this._sheet.Cells["A1"].Value = 10;
            this._sheet.Cells["A2"].Value = 11;
            this._sheet.Cells["A3"].Value = 12;
            this._sheet.Cells["A4"].Value = 13;
            this._sheet.Cells["A5"].Value = 14;
            this._sheet.Cells["A6"].Value = 15;
            this._sheet.Cells["B1"].Formula = "COUNTIF(A1:A6,\">13\")";
            this._sheet.Calculate();
            Assert.AreEqual(2d, this._sheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void CountIfShouldCalculateRangeCriteria()
        {
            this._sheet.Cells["A1"].Value = 10;
            this._sheet.Cells["A2"].Value = 11;
            this._sheet.Cells["A3"].Value = 12;
            this._sheet.Cells["A4"].Value = 12;
            this._sheet.Cells["A5"].Value = 14;
            this._sheet.Cells["A6"].Value = 15;

            this._sheet.Cells["B1"].Value = 10;
            this._sheet.Cells["B2"].Value = 12;

            this._sheet.Cells["C1"].Formula = "SUM(COUNTIF(A1:A6,B1:B2))";
            this._sheet.Calculate();
            Assert.AreEqual(3d, this._sheet.Cells["C1"].Value);
        }
    }
}
