using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml.Style;

namespace EPPlusTest.Core.Range;

[TestClass]
public class RangeRichTextTests : TestBase
{
    static ExcelPackage _pck;
    static ExcelWorksheet _ws;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        _pck = OpenPackage("RangeRichText.xlsx", true);
        _ws = _pck.Workbook.Worksheets.Add("Richtext");
    }

    [ClassCleanup]
    public static void Cleanup() => SaveAndCleanup(_pck);

    [TestMethod]
    public void AddThreeParagraphsAndValidate()
    {
        ExcelRange? r = _ws.Cells["A1"];
        ExcelRichText? r1 = r.RichText.Add("Line1\n");
        r1.PreserveSpace = true;
        ExcelRichText? r2 = r.RichText.Add("Line2\n");
        r2.PreserveSpace = true;
        ExcelRichText? r3 = r.RichText.Add("Line3");
        r3.PreserveSpace = true;
        r3.Bold = true;
        r3.Italic = true;
        r3.Size = 19.5F;

        Assert.AreEqual("Line1\nLine2\nLine3", r.Text);
        Assert.AreEqual("Line1\nLine2\nLine3", r.RichText.Text);
    }

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void AddEmptyStringShouldThrowArgumentException() => _ = _ws.Cells["D1"].RichText.Add(null);

    [TestMethod]
    [ExpectedException(typeof(ArgumentException))]
    public void AddNullShouldThrowArgumentException() => _ = _ws.Cells["D1"].RichText.Add(null);

    [TestMethod]
    [ExpectedException(typeof(InvalidOperationException))]
    public void SettingNameToEmptyStringShouldThrowInvalidOperationException()
    {
        _ = _ws.Cells["D1"].RichText.Add(" ");
        _ws.Cells["D1"].RichText[0].Text = null;
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidOperationException))]
    public void SettingTextToNullShouldThrowInvalidOperationException()
    {
        _ = _ws.Cells["D1"].RichText.Add(" ");
        _ws.Cells["D1"].RichText[0].Text = null;
    }

    [TestMethod]
    public void SettingRichTextTextToNullShouldClearRichText()
    {
        _ = _ws.Cells["D1"].RichText.Add(" ");
        Assert.IsTrue(_ws.Cells["D1"].IsRichText);
        _ws.Cells["D1"].RichText.Text = null;
        Assert.AreEqual(0, _ws.Cells["D1"].RichText.Count);
        Assert.IsFalse(_ws.Cells["D1"].IsRichText);
    }

    [TestMethod]
    public void SettingRichTextTextToEmptyStringShouldClearRichText()
    {
        _ = _ws.Cells["D1"].RichText.Add(" ");
        Assert.IsTrue(_ws.Cells["D1"].IsRichText);
        _ws.Cells["D1"].RichText.Text = null;
        Assert.AreEqual(0, _ws.Cells["D1"].RichText.Count);
        Assert.IsFalse(_ws.Cells["D1"].IsRichText);
    }

    [TestMethod]
    public void RemoveVerticalAlign()
    {
        ExcelRichText? p = _ws.Cells["G1"].RichText.Add("RemoveVerticalAlign");
        p.VerticalAlign = ExcelVerticalAlignmentFont.Baseline;
        p.VerticalAlign = ExcelVerticalAlignmentFont.None;
        Assert.AreEqual(p.VerticalAlign, ExcelVerticalAlignmentFont.None);
    }

    [TestMethod]
    public void ValidateIsRichTextValuesAndTexts()
    {
        using ExcelPackage? p1 = new ExcelPackage();
        ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("RichText");
        string? v = "Player's taunt success & you attack them";
        ws.Cells["A1"].Value = v;
        p1.Save();

        using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
        Assert.AreEqual(v, ws.Cells["A1"].Value);
        ws.Cells["A1"].IsRichText = true;
        Assert.AreEqual(v, ws.Cells["A1"].Value);
        Assert.AreEqual(v, ws.Cells["A1"].RichText.Text);
        ws.Cells["A1"].IsRichText = false;
        Assert.AreEqual(v, ws.Cells["A1"].Value);

        p2.Save();
    }

    [TestMethod]
    public void ValidateRichTextOverwriteByArray()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("RichTextOverwriteArray");

        for (int row = 1; row < 10; row++)
        {
            for (int col = 1; col < 10; col++)
            {
                ws.Cells[row, col].RichText.Text = $"Cell {ExcelCellBase.GetAddress(row, col)}";
            }
        }

        object[,]? array = new object[,] { { "Overwrite cell 1-1", "Overwrite cell 1-2" }, { "Overwrite cell 2-1", "Overwrite cell 2-2" } };

        ws.Cells["C3:D4"].Value = array;

        Assert.IsTrue(ws.Cells["C2"].IsRichText);
        Assert.IsTrue(ws.Cells["E3"].IsRichText);
        Assert.IsTrue(ws.Cells["A4"].IsRichText);
        Assert.IsTrue(ws.Cells["C5"].IsRichText);

        Assert.IsFalse(ws.Cells["C3"].IsRichText);
        Assert.IsFalse(ws.Cells["D4"].IsRichText);

        Assert.AreEqual("Cell C2", ws.Cells["C2"].Value);
        Assert.AreEqual("Overwrite cell 1-1", ws.Cells["C3"].Value);
        Assert.AreEqual("Overwrite cell 2-2", ws.Cells["D4"].Value);
        Assert.AreEqual("Cell A4", ws.Cells["A4"].Value);
        Assert.AreEqual("Cell D5", ws.Cells["D5"].Value);
    }

    [TestMethod]
    public void IsRichTextShouldKeepValues()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("IsRichTextKeepValues");
        CultureInfo? ci = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        ws.Cells["A1"].Value = "Cell A1";
        ws.Cells["B1"].Value = "Cell B1";
        ws.Cells["A2"].Value = 2;
        ws.Cells["B2"].Value = 3.2;

        ws.Cells["A1:B2"].IsRichText = true;

        Assert.IsTrue(ws.Cells["A1"].IsRichText);
        Assert.IsTrue(ws.Cells["B1"].IsRichText);
        Assert.IsTrue(ws.Cells["A2"].IsRichText);
        Assert.IsTrue(ws.Cells["B2"].IsRichText);

        Assert.AreEqual("Cell A1", ws.Cells["A1"].Value);
        Assert.AreEqual("Cell B1", ws.Cells["B1"].Value);
        Assert.AreEqual("2", ws.Cells["A2"].Value);
        Assert.AreEqual("3.2", ws.Cells["B2"].Value);

        ws.Cells["A1:B2"].IsRichText = false;

        Assert.AreEqual("Cell A1", ws.Cells["A1"].Value);
        Assert.AreEqual("Cell B1", ws.Cells["B1"].Value);
        Assert.AreEqual("2", ws.Cells["A2"].Value);
        Assert.AreEqual("3.2", ws.Cells["B2"].Value);

        Thread.CurrentThread.CurrentCulture = ci;
    }

    [TestMethod]
    public void ValidateRichText_TextIsReflectedOnRemove()
    {
        ExcelPackage? package = new ExcelPackage();
        _ = package.Workbook.Worksheets.Add("Test");
        ExcelRange? range = package.Workbook.Worksheets[0].Cells[1, 1];
        ExcelRichText? first = range.RichText.Add("1");
        ExcelRichText? second = range.RichText.Add("2");
        Assert.IsTrue(first != null);
        Assert.IsTrue(second != null);
        Assert.IsTrue(range.IsRichText);
        Assert.AreEqual(2, range.RichText.Count);
        Assert.AreEqual("12", range.Text);
        range.RichText.Remove(second);
        Assert.AreEqual(1, range.RichText.Count);
        Assert.AreEqual("1", range.Text); // FAILS as "12"
    }
}