using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System;
using System.IO;

namespace EPPlusTest.Core.Range;

[TestClass]
public class NamedRangeTests
{
    [TestMethod]
    public void IsValidName()
    {
        Assert.IsFalse(ExcelAddressUtil.IsValidName("123sa")); //invalid start char 
        Assert.IsFalse(ExcelAddressUtil.IsValidName("*d")); //invalid start char
        Assert.IsFalse(ExcelAddressUtil.IsValidName("\t")); //invalid start char
        Assert.IsFalse(ExcelAddressUtil.IsValidName("\\t")); //Backslash at least three chars
        Assert.IsFalse(ExcelAddressUtil.IsValidName("A+1")); //invalid char
        Assert.IsFalse(ExcelAddressUtil.IsValidName("A%we")); //Address invalid
        Assert.IsFalse(ExcelAddressUtil.IsValidName("BB73")); //Address invalid
        Assert.IsTrue(ExcelAddressUtil.IsValidName("BBBB75")); //Valid
        Assert.IsTrue(ExcelAddressUtil.IsValidName("BB1500005")); //Valid
    }

    [TestMethod]
    public void NamedRangeMovesDownIfRowInsertedAbove()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NEW");
        ExcelRange? range = sheet.Cells[2, 1, 3, 3];
        ExcelNamedRange? namedRange = sheet.Names.Add("NewNamedRange", range);

        sheet.InsertRow(1, 1);

        Assert.AreEqual("NEW!A3:C4", namedRange.FullAddress);
    }

    [TestMethod]
    public void NamedRangeDoesNotChangeIfRowInsertedBelow()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NEW");
        ExcelRange? range = sheet.Cells[2, 1, 3, 3];
        ExcelNamedRange? namedRange = sheet.Names.Add("NewNamedRange", range);

        sheet.InsertRow(4, 1);

        Assert.AreEqual("A2:C3", namedRange.Address);
    }

    [TestMethod]
    public void NamedRangeExpandsDownIfRowInsertedWithin()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NEW");
        ExcelRange? range = sheet.Cells[2, 1, 3, 3];
        ExcelNamedRange? namedRange = sheet.Names.Add("NewNamedRange", range);

        sheet.InsertRow(3, 1);

        Assert.AreEqual("NEW!A2:C4", namedRange.FullAddress);
    }

    [TestMethod]
    public void NamedRangeMovesRightIfColInsertedBefore()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NEW");
        ExcelRange? range = sheet.Cells[2, 2, 3, 4];
        ExcelNamedRange? namedRange = sheet.Names.Add("NewNamedRange", range);

        sheet.InsertColumn(1, 1);

        Assert.AreEqual("NEW!C2:E3", namedRange.FullAddress);
    }

    [TestMethod]
    public void NamedRangeUnchangedIfColInsertedAfter()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NEW");
        ExcelRange? range = sheet.Cells[2, 2, 3, 4];
        ExcelNamedRange? namedRange = sheet.Names.Add("NewNamedRange", range);

        sheet.InsertColumn(5, 1);

        Assert.AreEqual("B2:D3", namedRange.Address);
    }

    [TestMethod]
    public void NamedRangeExpandsToRightIfColInsertedWithin()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NEW");
        ExcelRange? range = sheet.Cells[2, 2, 3, 4];
        ExcelNamedRange? namedRange = sheet.Names.Add("NewNamedRange", range);

        sheet.InsertColumn(5, 1);

        Assert.AreEqual("B2:D3", namedRange.Address);
    }

    [TestMethod]
    public void NamedRangeWithWorkbookScopeIsMovedDownIfRowInsertedAbove()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorkbook? workbook = package.Workbook;
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NEW");
        ExcelRange? range = sheet.Cells[2, 1, 3, 3];
        ExcelNamedRange? namedRange = workbook.Names.Add("NewNamedRange", range);

        sheet.InsertRow(1, 1);

        Assert.AreEqual("NEW!A3:C4", namedRange.FullAddress);
    }

    [TestMethod]
    public void NamedRangeWithWorkbookScopeIsMovedRightIfColInsertedBefore()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorkbook? workbook = package.Workbook;
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NEW");
        ExcelRange? range = sheet.Cells[2, 2, 3, 3];
        ExcelNamedRange? namedRange = workbook.Names.Add("NewNamedRange", range);

        sheet.InsertColumn(1, 1);

        Assert.AreEqual("NEW!C2:D3", namedRange.FullAddress);
    }

    [TestMethod]
    public void NamedRangeIsUnchangedForOutOfScopeSheet()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorkbook? workbook = package.Workbook;
        ExcelWorksheet? sheet1 = package.Workbook.Worksheets.Add("NEW");
        ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("NEW2");
        ExcelRange? range = sheet2.Cells[2, 2, 3, 3];
        ExcelNamedRange? namedRange = workbook.Names.Add("NewNamedRange", range);

        sheet1.InsertColumn(1, 1);

        Assert.AreEqual("B2:C3", namedRange.Address);
    }

    [TestMethod]
    public void NamedRangeIsEqual()
    {
        using ExcelPackage? p1 = new ExcelPackage();
        using ExcelPackage? p2 = new ExcelPackage();
        ExcelWorksheet? ws1 = p1.Workbook.Worksheets.Add("sheet1");
        _ = p1.Workbook.Worksheets.Add("sheet2");

        ExcelWorksheet? ws1_p2 = p2.Workbook.Worksheets.Add("sheet1");

        ExcelNamedRange? wbName1 = p1.Workbook.Names.Add("Name1", ws1.Cells["sheet1!A1"]);
        ExcelNamedRange? wsName1 = ws1.Names.Add("Name1", ws1.Cells["A1"]);
        ExcelNamedRange? wsName2 = ws1.Names.Add("Name2", ws1.Cells["A1"]);

        ExcelNamedRange? wsName1_p2 = ws1_p2.Names.Add("Name1", ws1_p2.Cells["A1"]);

        //Assert
        Assert.IsTrue(wbName1.Equals(wbName1));
        Assert.IsTrue(wsName1.Equals(wsName1));

        Assert.IsFalse(wsName1.Equals(wbName1));
        Assert.IsFalse(wbName1.Equals(wsName2));
        Assert.IsFalse(wsName1.Equals(wsName1_p2));
    }

    [TestMethod]
    public void WorkbookNamedRange_ShouldRetain_FixedAddress()
    {
        using MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? package = new ExcelPackage(ms))
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            _ = package.Workbook.Names.Add("MyName", sheet.Cells["$A$1:$A$3"]);
            package.Save();
        }

        ms.Position = 0;
        using ExcelPackage? package2 = new ExcelPackage(ms);
        string? nameAddress = package2.Workbook.Names["MyName"].ToInternalAddress().Address;
        Assert.AreEqual("test!$A$1:$A$3", nameAddress);
    }

    [TestMethod]
    public void WorksheetNamedRange_ShouldRetain_FixedAddress()
    {
        using MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? package = new ExcelPackage(ms))
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            _ = sheet.Names.Add("MyName", sheet.Cells["$A$1:$A$3"]);
            package.Save();
        }

        ms.Position = 0;
        using ExcelPackage? package2 = new ExcelPackage(ms);
        string? nameAddress = package2.Workbook.Worksheets["test"].Names["MyName"].ToInternalAddress().Address;
        Assert.AreEqual("test!$A$1:$A$3", nameAddress);
    }

    [TestMethod]
    public void WorkbookNamedRange_ShouldRetainRelativeAddress_WhenIsRelativeIsTrue()
    {
        using MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? package = new ExcelPackage(ms))
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            _ = package.Workbook.Names.Add("MyName", sheet.Cells["A1:A3"], true);
            package.Save();
        }

        ms.Position = 0;
        using ExcelPackage? package2 = new ExcelPackage(ms);
        string? nameAddress = package2.Workbook.Names["MyName"].ToInternalAddress().Address;
        Assert.AreEqual("test!A1:A3", nameAddress);
    }

    [TestMethod]
    public void WorksheetNamedRange_ShouldRetainRelativeAddress_WhenIsRelativeIsTrue()
    {
        using MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? package = new ExcelPackage(ms))
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            _ = sheet.Names.Add("MyName", sheet.Cells["A1:A3"], true);
            package.Save();
        }

        ms.Position = 0;
        using ExcelPackage? package2 = new ExcelPackage(ms);
        string? nameAddress = package2.Workbook.Worksheets["test"].Names["MyName"].ToInternalAddress().Address;
        Assert.AreEqual("test!A1:A3", nameAddress);
    }

    [TestMethod]
    public void WorkbookNamedRange_ShouldNotRetainRelativeAddress_WhenIsRelativeIsFalse()
    {
        using MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? package = new ExcelPackage(ms))
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            _ = package.Workbook.Names.Add("MyName", sheet.Cells["A1:A3"], false);
            package.Save();
        }

        ms.Position = 0;
        using ExcelPackage? package2 = new ExcelPackage(ms);
        string? nameAddress = package2.Workbook.Names["MyName"].ToInternalAddress().Address;
        Assert.AreEqual("test!$A$1:$A$3", nameAddress);
    }

    [TestMethod]
    public void WorksheetNamedRange_ShouldNotRetainRelativeAddress_WhenIsRelativeIsFalse()
    {
        using MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? package = new ExcelPackage(ms))
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            _ = sheet.Names.Add("MyName", sheet.Cells["A1:A3"], false);
            package.Save();
        }

        ms.Position = 0;
        using ExcelPackage? package2 = new ExcelPackage(ms);
        string? nameAddress = package2.Workbook.Worksheets["test"].Names["MyName"].ToInternalAddress().Address;
        Assert.AreEqual("test!$A$1:$A$3", nameAddress);
    }

    [TestMethod]
    public void WorkbookNamedRange_ShouldAlwaysSetFixedAddress_WhenNotLoadingFromFile()
    {
        using MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? package = new ExcelPackage(ms))
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            _ = package.Workbook.Names.Add("MyName", sheet.Cells["A1:A3"]);
            package.Save();
        }

        ms.Position = 0;
        using ExcelPackage? package2 = new ExcelPackage(ms);
        string? nameAddress = package2.Workbook.Names["MyName"].ToInternalAddress().Address;
        Assert.AreEqual("test!$A$1:$A$3", nameAddress);
    }

    [TestMethod]
    public void WorksheetNamedRange_ShouldAlwaysSetFixedAddress_WhenNotLoadingFromFile()
    {
        using MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? package = new ExcelPackage(ms))
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
            _ = sheet.Names.Add("MyName", sheet.Cells["A1:A3"]);
            package.Save();
        }

        ms.Position = 0;
        using ExcelPackage? package2 = new ExcelPackage(ms);
        string? nameAddress = package2.Workbook.Worksheets["test"].Names["MyName"].ToInternalAddress().Address;
        Assert.AreEqual("test!$A$1:$A$3", nameAddress);
    }
}