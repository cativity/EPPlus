using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.ThreadedComments;

namespace EPPlusTest.ThreadedComments;

[TestClass]
public class ThreadedCommentsCopyTests : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("ThreadedCommentCopy.xlsx", true);
    }

    [ClassCleanup]
    public static void Cleanup()
    {
        SaveAndCleanup(_pck);
    }

    [TestMethod]
    public void ShouldCopyThreadedCommentWithinSheet()
    {
        ExcelWorksheet? sheet = _pck.Workbook.Worksheets.Add("WithinSheet");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentPerson? person2 = sheet.ThreadedComments.Persons.Add("Jane Doe");
        ExcelThreadedCommentThread? thread = sheet.Cells["A1"].AddThreadedComment();
        ExcelThreadedComment? c1 = thread.AddComment(person2.Id, "Hello");
        ExcelThreadedComment? c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

        sheet.Cells[1, 1].Copy(sheet.Cells["A3"]);
        thread = sheet.Cells[3, 1].ThreadedComment;

        Assert.AreEqual(2, thread.Comments.Count);
        Assert.AreEqual("A3", thread.Comments[0].Ref);
        Assert.AreEqual("A3", thread.Comments[1].Ref);
        Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
        Assert.AreEqual(1, thread.Comments[1].Mentions.Count());
    }

    [TestMethod]
    public void ShouldCopyThreadedCommentToNewSheet()
    {
        ExcelWorksheet? sheet = _pck.Workbook.Worksheets.Add("NewSheet_Source");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentPerson? person2 = sheet.ThreadedComments.Persons.Add("Jane Doe");
        ExcelThreadedCommentThread? thread = sheet.Cells["A1"].AddThreadedComment();
        ExcelThreadedComment? c1 = thread.AddComment(person2.Id, "Hello");
        ExcelThreadedComment? c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

        ExcelWorksheet? sheet2 = _pck.Workbook.Worksheets.Add("NewSheet_Dest");
        sheet.Cells[1, 1].Copy(sheet2.Cells["A3"]);
        thread = sheet2.Cells[3, 1].ThreadedComment;

        Assert.AreEqual(2, thread.Comments.Count);
        Assert.AreEqual("A3", thread.Comments[0].Ref);
        Assert.AreEqual("A3", thread.Comments[1].Ref);
        Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
        Assert.AreEqual(1, thread.Comments[1].Mentions.Count());
    }

    [TestMethod]
    public void ShouldCopyThreadedCommentToNewPackage()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? sheet = p.Workbook.Worksheets.Add("test");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentPerson? person2 = sheet.ThreadedComments.Persons.Add("Jane Doe");
        ExcelThreadedCommentThread? thread = sheet.Cells["A1"].AddThreadedComment();
        ExcelThreadedComment? c1 = thread.AddComment(person2.Id, "Hello");
        ExcelThreadedComment? c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);
        using ExcelPackage? pck2 = new ExcelPackage();
        ExcelWorksheet? sheet2 = pck2.Workbook.Worksheets.Add("test2");
        sheet.Cells[1, 1].Copy(sheet2.Cells["A3"]);
        thread = sheet2.Cells[3, 1].ThreadedComment;

        Assert.AreEqual(2, thread.Comments.Count);
        Assert.AreEqual("A3", thread.Comments[0].Ref);
        Assert.AreEqual("A3", thread.Comments[1].Ref);
        Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
        Assert.AreEqual(1, thread.Comments[1].Mentions.Count());
        SaveWorkbook("ThreadedCommentCopy_NewPackage.xlsx", pck2);
    }

    [TestMethod]
    public void ShouldCopyWorksheetWithThreadedComment()
    {
        ExcelWorksheet? sheetToCopy = _pck.Workbook.Worksheets.Add("WorksheetCopy_Source");
        ExcelThreadedCommentPerson? person = sheetToCopy.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentPerson? person2 = sheetToCopy.ThreadedComments.Persons.Add("Jane Doe");
        ExcelThreadedCommentThread? thread = sheetToCopy.Cells["A1"].AddThreadedComment();
        ExcelThreadedComment? c1 = thread.AddComment(person2.Id, "Hello");
        ExcelThreadedComment? c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

        ExcelWorksheet? copy = _pck.Workbook.Worksheets.Add("WorksheetCopy_Dest", sheetToCopy);
        thread = copy.Cells[1, 1].ThreadedComment;

        Assert.AreEqual(2, thread.Comments.Count);
        Assert.AreEqual("A1", thread.Comments[0].Ref);
        Assert.AreEqual("A1", thread.Comments[1].Ref);
        Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
        Assert.AreEqual(1, thread.Comments[1].Mentions.Count());
    }

    [TestMethod]
    public void ShouldCopyWorksheetWithThreadedCommentToNewPackage()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? sheetToCopy = p.Workbook.Worksheets.Add("WorksheetCopy_Source");
        ExcelThreadedCommentPerson? person = sheetToCopy.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentPerson? person2 = sheetToCopy.ThreadedComments.Persons.Add("Jane Doe");
        ExcelThreadedCommentThread? thread = sheetToCopy.Cells["A1"].AddThreadedComment();
        ExcelThreadedComment? c1 = thread.AddComment(person2.Id, "Hello");
        ExcelThreadedComment? c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

        using ExcelPackage? pck2 = new ExcelPackage();
        ExcelWorksheet? copy = pck2.Workbook.Worksheets.Add("WorksheetCopy_Desc", sheetToCopy);
        thread = copy.Cells[1, 1].ThreadedComment;

        Assert.AreEqual(2, thread.Comments.Count);
        Assert.AreEqual("A1", thread.Comments[0].Ref);
        Assert.AreEqual("A1", thread.Comments[1].Ref);
        Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
        Assert.AreEqual(1, thread.Comments[1].Mentions.Count());

        SaveWorkbook("ThreadedCommentWorksheetCopy_NewPackage.xlsx", pck2);
    }
}