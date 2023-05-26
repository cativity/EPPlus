using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.SaveFunctions;

[TestClass]
public class ToTextTests
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
        this._package.Dispose();
    }

    [TestMethod]
    public void ToTextTextDefault()
    {
        this._sheet.Cells["A1"].Value = "h1";
        this._sheet.Cells["B1"].Value = "h2";
        string? text = this._sheet.Cells["A1:B1"].ToText();
        Assert.AreEqual("h1,h2", text);
    }

    [TestMethod]
    public void ToTextTextMultilines()
    {
        this._sheet.Cells["A1"].Value = "h1";
        this._sheet.Cells["B1"].Value = "h2";
        this._sheet.Cells["A2"].Value = 1;
        this._sheet.Cells["B2"].Value = 2;
        string? text = this._sheet.Cells["A1:B2"].ToText();
        Assert.AreEqual("h1,h2" + Environment.NewLine + "1,2", text);
    }

    [TestMethod]
    public void ToTextTextTextQualifier()
    {
        this._sheet.Cells["A1"].Value = "h1";
        this._sheet.Cells["B1"].Value = "h2";
        this._sheet.Cells["A2"].Value = 1;
        this._sheet.Cells["B2"].Value = 2;
        ExcelOutputTextFormat? format = new ExcelOutputTextFormat
        {
            TextQualifier = '\''
        };
        string? text = this._sheet.Cells["A1:B2"].ToText(format);
        Assert.AreEqual("'h1','h2'" + Environment.NewLine + "1,2", text);
    }
    [TestMethod]
    public void ToTextTextQualifierWithNumericContaingSeparator()
    {
        this._sheet.Cells["A10"].Value = "h1";
        this._sheet.Cells["B10"].Value = "h2";
        this._sheet.Cells["A11"].Value = 1;
        this._sheet.Cells["B11"].Value = 2;
        this._sheet.Cells["A11:B11"].Style.Numberformat.Format = "#,##0.00";
        ExcelOutputTextFormat? format = new ExcelOutputTextFormat
        {
            TextQualifier = '\"',
            DecimalSeparator = ",",
            UseCellFormat = true,
            Culture = new System.Globalization.CultureInfo("sv-SE")
        };
        string? text = this._sheet.Cells["A10:B11"].ToText(format);
        Assert.AreEqual("\"h1\",\"h2\"" + Environment.NewLine + "\"1,00\",\"2,00\"", text);
    }

    [TestMethod]
    public void ToTextTextIgnoreHeaders()
    {
        this._sheet.Cells["A1"].Value = 1;
        this._sheet.Cells["B1"].Value = 2;
        ExcelOutputTextFormat? format = new ExcelOutputTextFormat
        {
            TextQualifier = '\'',
            FirstRowIsHeader = false
        };
        string? text = this._sheet.Cells["A1:B1"].ToText(format);
        Assert.AreEqual("1,2", text);
    }
}