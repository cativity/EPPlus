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
    public class IfErrorTests
    {
        [TestMethod]
        public void IfError_InnerFunctionReturningError()
        {
            using ExcelPackage? pck = new ExcelPackage();
            ExcelWorksheet? sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
            sheet1.Cells["C3"].Formula = "IFERROR(IF(NameDoesntExist=1,\"A\",\"B\"),\"error\")";

            sheet1.Calculate();

            Assert.IsFalse(sheet1.Cells["C3"].Value is ExcelErrorValue);
            Assert.AreEqual("error", sheet1.Cells["C3"].GetValue<string>());
        }
    }
}
