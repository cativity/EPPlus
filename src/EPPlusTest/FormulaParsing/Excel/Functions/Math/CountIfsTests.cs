﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class CountIfsTests
    {
        [TestMethod]
        public void CountIfsShouldNotCountNumericStringsAsNumbers()
        {
            using(ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = "123";
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,\">0\")";
                sheet.Calculate();
                object? val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(0d, val);
            }
        }

        [TestMethod]
        public void CountIfsShouldCountMatchingNumericValue()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = 123;
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,\">0\")";
                sheet.Calculate();
                object? val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(1d, val);
            }
        }
    }
}
