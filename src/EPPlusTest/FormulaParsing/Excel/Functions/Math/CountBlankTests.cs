using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class CountBlankTests
    {
        [TestMethod]
        public void ShouldCountEmptyCells()
        {
            using(ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Count_Blank");
                sheet.Cells["A1"].Value = "Test 1";
                sheet.Cells["A2"].Value = "Test 2";
                sheet.Cells["A3"].Value = "Test 3";

                sheet.Cells["B5"].Formula = "COUNTBLANK($B$1:B1)";
                sheet.Cells["B6"].Formula = "COUNTBLANK($B$1)";
                sheet.Cells["B7"].Formula = "COUNTBLANK(B1)";
                sheet.Cells["B8"].Formula = "COUNTBLANK($B$1:B3)";

                sheet.Calculate();

                object? b5 = sheet.Cells["B5"].Value;
                object? b6 = sheet.Cells["B6"].Value;
                object? b7 = sheet.Cells["B7"].Value;
                object? b8 = sheet.Cells["B8"].Value;

                Assert.AreEqual(1, b5);
                Assert.AreEqual(1, b6);
                Assert.AreEqual(1, b7);
                Assert.AreEqual(3, b8);

            }
        }
    }
}
