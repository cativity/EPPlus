﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class SumIfsTests
{
    [TestMethod]
    public void SumIfsShouldHandleSingleRange()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].Formula = "SUMIFS(H5,H5,\">0\",K5,\"> 0\")";
        sheet.Cells["H5"].Value = 1;
        sheet.Cells["K5"].Value = 1;
        sheet.Calculate();
        Assert.AreEqual(1d, sheet.Cells["A1"].Value);
    }
}