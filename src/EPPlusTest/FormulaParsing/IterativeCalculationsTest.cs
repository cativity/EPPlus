using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing;

[TestClass]
public class IterativeCalculationsTest
{
    private ExcelPackage _package;
    private ExcelWorksheet _sheet;

    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        this._sheet = this._package.Workbook.Worksheets.Add("test");
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._sheet = null;
        this._package.Dispose();
    }

    //[TestMethod]
    //public void A1andB1CircularRegShouldWork()
    //{
    //    _sheet.Cells["B1"].Value = 1;
    //    _sheet.Cells["A1"].Formula = "A2 + B1";
    //    _sheet.Cells["A2"].Formula = "A1 + B1";

    //    var options = new ExcelCalculationOption { AllowCircularReferences = true };
    //    _sheet.Calculate(options);

    //    Assert.AreEqual(1d, _sheet.Cells["A1"].Value);
    //    Assert.AreEqual(2d, _sheet.Cells["B1"].Value);
    //}

}