using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core;

[TestClass]
public class LoadSaveTests : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context) => _pck = OpenPackage("LoadSaveTest.xlsx", true);

    [ClassCleanup]
    public static void Cleanup() => SaveAndCleanup(_pck);

    [TestMethod]
    public void CheckCfLfIsRetained()
    {
        using ExcelPackage? p1 = new ExcelPackage();
        string? expected = "Line1\r\nLine2";
        ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("CrLf");
        ws.Cells["A1"].Value = expected;
        Assert.AreEqual(expected, ws.Cells["A1"].Value);

        ws.Save();

        using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
        ws = p1.Workbook.Worksheets["CrLf"];
        Assert.AreEqual(expected, ws.Cells["A1"].Value);
    }

    [TestMethod]
    public void SaveTwiceShouldNotCorruptPackage()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("SaveTwice");
        p.Workbook.Properties.Application = "EPPlus";
        ws.Cells["A1"].Value = "A1";
        p.Workbook.Properties.Title = "EPPlus";
        p.Save();
        long length = p.Stream.Length;
        byte[]? b = p.GetAsByteArray();
        Assert.AreEqual(length, b.Length);
    }

    [TestMethod]
    public async Task SaveTwiceShouldNotCorruptPackageAsync()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("SaveTwice");
        p.Workbook.Properties.Application = "EPPlus";
        ws.Cells["A1"].Value = "A1";
        p.Workbook.Properties.Title = "EPPlus";
        p.Save();
        long length = p.Stream.Length;
        byte[]? b = await p.GetAsByteArrayAsync();

        Assert.AreEqual(length, b.Length);
    }

    [TestMethod]
    public void ChartSheetShouldNotThrowException()
    {
        using ExcelPackage? p = new ExcelPackage();
        ExcelWorksheet? s1 = p.Workbook.Worksheets.Add("Table1");
        _ = p.Workbook.Worksheets.AddChart("Chart1", OfficeOpenXml.Drawing.Chart.eChartType.Area);
        _ = p.Workbook.Worksheets.Add("Table2");

        DataTable dt = new DataTable();

        dt.Columns.Add(new DataColumn("Title", typeof(string)));
        dt.Columns.Add(new DataColumn("Count", typeof(int)));

        DataRow? r1 = dt.NewRow();
        DataRow? r2 = dt.NewRow();
        DataRow? r3 = dt.NewRow();

        r1.ItemArray = new object[] { "Title", 20 };
        r2.ItemArray = new object[] { "Title", 20 };
        r3.ItemArray = new object[] { "Title", 20 };

        dt.Rows.Add(r1);
        dt.Rows.Add(r2);
        dt.Rows.Add(r3);

        _ = s1.Cells[1, 1, 3, 2].LoadFromDataTable(dt, true, null);
    }

    [TestMethod]
    public void VerifyInvalidXmlUnicodeChar()
    {
        string s1 = "String with \ufffe char";
        string s2 = "Second string with \uffff char";
        using ExcelPackage? p1 = new ExcelPackage();
        ExcelWorksheet? ws = p1.Workbook.Worksheets.Add("Sheet1");
        ws.SetValue(1, 1, s1);
        ws.SetValue(2, 1, s2);
        p1.Save();

        using ExcelPackage? p2 = new ExcelPackage(p1.Stream);
        Assert.AreEqual(s1, p2.Workbook.Worksheets[0].Cells["A1"].Value);
        Assert.AreEqual(s2, p2.Workbook.Worksheets[0].Cells["A2"].Value);
    }

    [TestMethod]
    public void LoadFromText_VerifyWithApostrophes()
    {
        string? textToLoad = "\"dog 1\"\"\"\"\"\"\",\"dog 2\"\"\"\"\"\"\",\"dog 3\"\"\"\"\"\"\"\r\n"
                             + "\"cat 1\",\"cat 2\",\"cat 3\"\"\"\"\"\r\n"
                             + "\"mouse 1\"\"\"\"\",\"mouse 2\"\"\"\"\",\"mouse 3\"\"\"\"\"";

        ExcelPackage? excelPackage = new ExcelPackage();
        ExcelWorksheet? ws = excelPackage.Workbook.Worksheets.Add("LoadFromText");
        _ = ws.Cells["B2"].LoadFromText(textToLoad, new ExcelTextFormat() { TextQualifier = '\"' });

        //Assert
        Assert.AreEqual("dog 1\"\"\"", ws.GetValue(2, 2));
        Assert.AreEqual("dog 2\"\"\"", ws.GetValue(2, 3));
        Assert.AreEqual("dog 3\"\"\"", ws.GetValue(2, 4));

        Assert.AreEqual("cat 1", ws.GetValue(3, 2));
        Assert.AreEqual("cat 2", ws.GetValue(3, 3));
        Assert.AreEqual("cat 3\"\"", ws.GetValue(3, 4));

        Assert.AreEqual("mouse 1\"\"", ws.GetValue(4, 2));
        Assert.AreEqual("mouse 2\"\"", ws.GetValue(4, 3));
        Assert.AreEqual("mouse 3\"\"", ws.GetValue(4, 4));
    }

    [TestMethod]
    public void SaveToText_VerifyWithApostrophes()
    {
        MemoryStream? ms = new MemoryStream();

        using (ExcelPackage? p = new ExcelPackage())
        {
            ExcelWorksheet? ws = p.Workbook.Worksheets.Add("SaveToText");
            ws.SetValue(2, 2, "dog 1\"\"\"");
            ws.SetValue(2, 3, "dog 2\"\"\"");
            ws.SetValue(2, 4, "dog 3\"\"\"");

            ws.SetValue(3, 2, "cat 1");
            ws.SetValue(3, 3, "cat 2");
            ws.SetValue(3, 4, "cat 3\"\"");

            ws.SetValue(4, 2, "mouse 1\"\"");
            ws.SetValue(4, 3, "mouse 2\"\"");
            ws.SetValue(4, 4, "mouse 3\"\"");

            ws.Cells["B2:D4"].SaveToText(ms, new ExcelOutputTextFormat() { TextQualifier = '\"' });
        }

        string? result;
        ms.Position = 0;

        using (StreamReader? reader = new StreamReader(ms))
        {
            result = reader.ReadToEnd();
        }

        //Assert
        string? expectedText = "\"dog 1\"\"\"\"\"\"\",\"dog 2\"\"\"\"\"\"\",\"dog 3\"\"\"\"\"\"\"\r\n"
                               + "\"cat 1\",\"cat 2\",\"cat 3\"\"\"\"\"\r\n"
                               + "\"mouse 1\"\"\"\"\",\"mouse 2\"\"\"\"\",\"mouse 3\"\"\"\"\"";

        Assert.AreEqual(expectedText, result);
    }

    private static ExcelPackage LoadFromText(FileInfo file, ExcelTextFormat format)
    {
        ExcelPackage? excelPackage = new ExcelPackage();
        ExcelWorksheet? sheet = excelPackage.Workbook.Worksheets.Add("bugs");
        _ = sheet.Cells.LoadFromText(file, format);

        return excelPackage;
    }
}