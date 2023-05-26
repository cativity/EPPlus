﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.ThreadedComments;

namespace EPPlusTest.ThreadedComments;

[TestClass]
public class ThreadedCommentsUnitTests
{

    [TestMethod]
    public void ShouldRemoveOnePerson()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentPerson? person2 = sheet.ThreadedComments.Persons.Add("John Does brother");
        Assert.AreEqual(2, package.Workbook.ThreadedCommentPersons.Count);
        package.Workbook.ThreadedCommentPersons.Remove(person2);
        Assert.AreEqual(1, package.Workbook.ThreadedCommentPersons.Count);
        Assert.AreEqual("John Doe", package.Workbook.ThreadedCommentPersons.First().DisplayName);
    }
    [TestMethod]
    public void ShouldAddOneComment()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        sheet.Cells["A1"].AddThreadedComment().AddComment(person.Id, "Hello");

        Assert.AreEqual(1, sheet.ThreadedComments.Count);
        Assert.AreEqual(1, sheet.Cells["A1"].ThreadedComment.Comments.Count);
    }

    [TestMethod]
    public void ShouldRemoveThread()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentThread? thread = sheet.Cells["A1"].AddThreadedComment();
        thread.AddComment(person.Id, "Hello");

        Assert.AreEqual(1, sheet.ThreadedComments.Count);
        Assert.AreEqual(1, sheet.Cells["A1"].ThreadedComment.Comments.Count);
        Assert.IsNotNull(sheet.Cells["A1"].Comment);

        sheet.ThreadedComments.Remove(thread);

        package.Save();

        Assert.IsNull(sheet.ThreadedComments["A1"]);
        Assert.IsNull(sheet.Comments["A1"]);
        Assert.IsNull(sheet.Cells["A1"].ThreadedComment);
        Assert.AreEqual(0, sheet.ThreadedComments.Count);
    }

    [TestMethod]
    public void ShouldRemoveOneComment()
    {
        using ExcelPackage? package = new ExcelPackage();


        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentThread? thread = sheet.Cells["A1"].AddThreadedComment();
        ExcelThreadedComment? c1 = thread.AddComment(person.Id, "Hello");
        ExcelThreadedComment? c2 = thread.AddComment(person.Id, "Hello again");
        Assert.AreEqual(2, thread.Comments.Count);

        bool rmResult = thread.Remove(c2);
        Assert.IsTrue(rmResult);
        Assert.AreEqual(1, thread.Comments.Count);
    }

    [TestMethod]
    public void ShouldAddMention()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentPerson? person2 = sheet.ThreadedComments.Persons.Add("Jane Doe");
        ExcelThreadedCommentThread? thread = sheet.Cells["A1"].AddThreadedComment();
        ExcelThreadedComment? c1 = thread.AddComment(person2.Id, "Hello");
        ExcelThreadedComment? c2 = thread.AddComment(person.Id, "Hello {0}", person2);

        Assert.AreEqual(2, thread.Comments.Count);
        Assert.AreEqual("Hello @Jane Doe", c2.Text);
        Assert.AreEqual(1, c2.Mentions.Count());
    }

    [TestMethod]
    public void ShouldRemoveMention()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        ExcelThreadedCommentPerson? person = sheet.ThreadedComments.Persons.Add("John Doe");
        ExcelThreadedCommentPerson? person2 = sheet.ThreadedComments.Persons.Add("Jane Doe");
        ExcelThreadedCommentThread? thread = sheet.Cells["A1"].AddThreadedComment();
        ExcelThreadedComment? c1 = thread.AddComment(person2.Id, "Hello");
        ExcelThreadedComment? c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

        Assert.AreEqual(2, thread.Comments.Count);
        Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
        Assert.AreEqual(1, c2.Mentions.Count());

        c2.EditText("Hello");
        Assert.AreEqual(0, c2.Mentions.Count());

        //package.SaveAs(new FileInfo("c:\\Temp\\JohnDoe.xlsx"));
    }
}