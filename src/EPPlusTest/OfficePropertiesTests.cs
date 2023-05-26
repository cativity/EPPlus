using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest;

[TestClass]
public class OfficePropertiesTests
{
        
    [TestMethod]
    public void ValidateLong()
    {
        using ExcelPackage? pck = new ExcelPackage();
        long ticks = DateTime.Now.Ticks;
        pck.Workbook.Properties.SetCustomPropertyValue("Timestamp", ticks);
        pck.Workbook.Worksheets.Add("Test");

        pck.Save();

        using ExcelPackage? pck2 = new ExcelPackage(pck.Stream);
        Assert.AreEqual((double)ticks, pck.Workbook.Properties.GetCustomPropertyValue("Timestamp"));
    }
    [TestMethod]
    public void ValidateCaseInsensitiveCustomProperties()
    {
        using ExcelPackage? p = new ExcelPackage();
        p.Workbook.Worksheets.Add("CustomProperties");
        p.Workbook.Properties.SetCustomPropertyValue("Foo", "Bar");
        p.Workbook.Properties.SetCustomPropertyValue("fOO", "bAR");

        Assert.AreEqual("bAR", p.Workbook.Properties.GetCustomPropertyValue("fOo"));
    }
    [TestMethod]
    public void ValidateCaseInsensitiveCustomProperties_Loading()
    {
        ExcelPackage? p = new ExcelPackage();
        p.Workbook.Worksheets.Add("CustomProperties");
        p.Workbook.Properties.SetCustomPropertyValue("fOO", "bAR");
        p.Workbook.Properties.SetCustomPropertyValue("Foo", "Bar");

        p.Save();

        ExcelPackage? p2 = new ExcelPackage(p.Stream);

        Assert.AreEqual("Bar", p2.Workbook.Properties.GetCustomPropertyValue("fOo"));

        p.Dispose();
        p2.Dispose();
    }
}