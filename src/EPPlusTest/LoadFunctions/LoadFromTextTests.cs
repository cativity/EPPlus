using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions;

[TestClass]
public class LoadFromTextTests
{
    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        this._worksheet = this._package.Workbook.Worksheets.Add("test");
        this._lines = new StringBuilder();
        this._format = new ExcelTextFormat();
    }

    [TestCleanup]
    public void Cleanup() => this._package.Dispose();

    private ExcelPackage _package;
    private ExcelWorksheet _worksheet;
    private StringBuilder _lines;
    private ExcelTextFormat _format;

    private void AddLine(string s) => _ = this._lines.AppendLine(s);

    [TestMethod]
    public void ShouldLoadCsvFormat()
    {
        this.AddLine("a,b,c");
        _ = this._worksheet.Cells["A1"].LoadFromText(this._lines.ToString());
        Assert.AreEqual("a", this._worksheet.Cells["A1"].Value);
    }

    [TestMethod]
    public void ShouldLoadCsvFormatWithDelimiter()
    {
        this.AddLine("a;b;c");
        this.AddLine("d;e;f");
        this._format.Delimiter = ';';
        _ = this._worksheet.Cells["A1"].LoadFromText(this._lines.ToString(), this._format);
        Assert.AreEqual("a", this._worksheet.Cells["A1"].Value);
        Assert.AreEqual("d", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void ShouldUseTypesFromFormat()
    {
        this.AddLine("a;2;5%");
        this.AddLine("d;3;8%");
        this._format.Delimiter = ';';
        this._format.DataTypes = new eDataTypes[] { eDataTypes.String, eDataTypes.Number, eDataTypes.Percent };
        _ = this._worksheet.Cells["A1"].LoadFromText(this._lines.ToString(), this._format);
        Assert.AreEqual("a", this._worksheet.Cells["A1"].Value);
        Assert.AreEqual(2d, this._worksheet.Cells["B1"].Value);
        Assert.AreEqual(3d, this._worksheet.Cells["B2"].Value);
        Assert.AreEqual(0.05, this._worksheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldUseHeadersFromFirstRow()
    {
        this.AddLine("Height 1,Width");
        this.AddLine("1,2");
        _ = this._worksheet.Cells["A1"].LoadFromText(this._lines.ToString(), this._format, TableStyles.None, true);
        Assert.AreEqual("Height 1", this._worksheet.Cells["A1"].Value);
        Assert.AreEqual(1d, this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void ShouldUseTextQualifier()
    {
        this.AddLine("'Look, a bird!',2");
        this.AddLine("'One apple, one orange',3");
        this._format.TextQualifier = '\'';
        _ = this._worksheet.Cells["A1"].LoadFromText(this._lines.ToString(), this._format);
        Assert.AreEqual("Look, a bird!", this._worksheet.Cells["A1"].Value);
        Assert.AreEqual("One apple, one orange", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void ShouldReturnRange()
    {
        this.AddLine("a,b,c");
        ExcelRangeBase? r = this._worksheet.Cells["A1"].LoadFromText(this._lines.ToString());
        Assert.AreEqual("A1:C2", r.FirstAddress);
    }

    [TestMethod]
    public void VerifyOneLineWithTextQualifier()
    {
        this.AddLine("\"a\",\"\"\"\", \"\"\"\"");
        _ = this._worksheet.Cells["A1"].LoadFromText(this._lines.ToString(), new ExcelTextFormat { TextQualifier = '\"' });
        Assert.AreEqual("a", this._worksheet.Cells[1, 1].Value);
        Assert.AreEqual("\"", this._worksheet.Cells[1, 2].Value);
        Assert.AreEqual("\"", this._worksheet.Cells[1, 3].Value);
    }

    [TestMethod]
    public void VerifyMultiLineWithTextQualifier()
    {
        this.AddLine("\"a\",b, \"c\"\"\"");
        this.AddLine("a,\"b\", \"c\"\"\r\n\"\"\"");
        this.AddLine("a,\"b\", \"c\"\"\"\"\"");
        this.AddLine("\"d\",e, \"\"");
        this.AddLine("\"\",, \"\"");

        _ = this._worksheet.Cells["A1"].LoadFromText(this._lines.ToString(), new ExcelTextFormat { TextQualifier = '\"' });
        Assert.AreEqual("a", this._worksheet.Cells[1, 1].Value);
        Assert.AreEqual("b", this._worksheet.Cells[1, 2].Value);
        Assert.AreEqual("c\"", this._worksheet.Cells[1, 3].Value);

        Assert.AreEqual("a", this._worksheet.Cells[2, 1].Value);
        Assert.AreEqual("b", this._worksheet.Cells[2, 2].Value);
        Assert.AreEqual("c\"\r\n\"", this._worksheet.Cells[2, 3].Value);

        Assert.AreEqual("a", this._worksheet.Cells[3, 1].Value);
        Assert.AreEqual("b", this._worksheet.Cells[3, 2].Value);
        Assert.AreEqual("c\"\"", this._worksheet.Cells[3, 3].Value);

        Assert.AreEqual("d", this._worksheet.Cells[4, 1].Value);
        Assert.AreEqual("e", this._worksheet.Cells[4, 2].Value);
        Assert.IsNull(this._worksheet.Cells[4, 3].Value);

        Assert.IsNull(this._worksheet.Cells[5, 1].Value);
        Assert.IsNull(this._worksheet.Cells[5, 2].Value);
        Assert.IsNull(this._worksheet.Cells[5, 3].Value);
    }
}