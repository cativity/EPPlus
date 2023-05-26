using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions;

[TestClass]
public class DateDifTests
{
    [TestMethod]
    public void ShouldHandleYearDiff()
    {
        CultureInfo? culture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        using (ExcelPackage? package = new ExcelPackage())
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
                
            sheet.Cells["A1"].Value = "1/1/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"Y\")";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["C1"].Value);

            sheet.Cells["A1"].Value = "1/4/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"Y\")";
            sheet.Calculate();
            Assert.AreEqual(1d, sheet.Cells["C1"].Value);
        }
        Thread.CurrentThread.CurrentCulture = culture;
    }

    [TestMethod]
    public void ShouldHandleMonthDiff()
    {
        CultureInfo? culture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        using (ExcelPackage? package = new ExcelPackage())
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = "1/1/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"M\")";
            sheet.Calculate();
            Assert.AreEqual(24d, sheet.Cells["C1"].Value);

            sheet.Cells["A1"].Value = "4/2/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"M\")";
            sheet.Calculate();
            Assert.AreEqual(20d, sheet.Cells["C1"].Value);
        }
        Thread.CurrentThread.CurrentCulture = culture;
    }

    [TestMethod]
    public void ShouldHandleTotalDays()
    {
        CultureInfo? culture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        using (ExcelPackage? package = new ExcelPackage())
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = "1/1/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"d\")";
            sheet.Calculate();
            Assert.AreEqual(730d, sheet.Cells["C1"].Value);

            sheet.Cells["A1"].Value = "4/2/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"d\")";
            sheet.Calculate();
            Assert.AreEqual(639d, sheet.Cells["C1"].Value);
        }
        Thread.CurrentThread.CurrentCulture = culture;
    }

    [TestMethod]
    public void ShouldHandleTotalDaysYm()
    {
        CultureInfo? culture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        using (ExcelPackage? package = new ExcelPackage())
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = "1/1/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"ym\")";
            sheet.Calculate();
            Assert.AreEqual(0d, sheet.Cells["C1"].Value);

            sheet.Cells["A1"].Value = "4/2/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 4, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"ym\")";
            sheet.Calculate();
            Assert.AreEqual(11d, sheet.Cells["C1"].Value);
        }
        Thread.CurrentThread.CurrentCulture = culture;
    }

    [TestMethod]
    public void ShouldHandleTotalDaysYd()
    {
        CultureInfo? culture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        using (ExcelPackage? package = new ExcelPackage())
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = "1/1/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"yd\")";
            sheet.Calculate();
            Assert.AreEqual(0d, sheet.Cells["C1"].Value);

            sheet.Cells["A1"].Value = "4/2/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"yd\")";
            sheet.Calculate();
            Assert.AreEqual(274d, sheet.Cells["C1"].Value);
        }
        Thread.CurrentThread.CurrentCulture = culture;
    }

    [TestMethod]
    public void ShouldHandleTotalDaysMd()
    {
        CultureInfo? culture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        using (ExcelPackage? package = new ExcelPackage())
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

            sheet.Cells["A1"].Value = "1/5/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 1, 6).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"md\")";
            sheet.Calculate();
            Assert.AreEqual(1d, sheet.Cells["C1"].Value);

            sheet.Cells["A1"].Value = "4/2/2001";
            sheet.Cells["B1"].Value = new DateTime(2003, 4, 1).ToOADate();
            sheet.Cells["C1"].Formula = "DATEDIF(A1,B1,\"md\")";
            sheet.Calculate();
            Assert.AreEqual(29d, sheet.Cells["C1"].Value);
        }
        Thread.CurrentThread.CurrentCulture = culture;
    }
}