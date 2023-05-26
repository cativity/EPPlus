using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering
{
    [TestClass]
    public class ComplexTests
    {
        [TestMethod]
        public void ComplexShouldReturnCorrectResult()
        {
            string? comma = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                
                sheet.Cells["A1"].Formula = "COMPLEX(5,2)";
                sheet.Calculate();
                object? result = sheet.Cells["A1"].Value;
                Assert.AreEqual("5+2i", result);

                sheet.Cells["A1"].Formula = "COMPLEX(5,-2.5, \"j\")";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual($"5-2{comma}5j", result);
            }
        }
    }
}
