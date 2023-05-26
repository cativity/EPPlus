using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;


namespace OfficeOpenXml.FormulaParsing;

[TestClass]
public class SharedFormulasTest
{
    [TestMethod]
    public void SharedFormulasShouldNotEffectFullColumn()
    {
        ExcelWorksheet.Formulas? f=new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default) { Index = 0, Formula = "SUM(C:D)", Address = "A1:B2", StartRow = 1, StartCol = 1 };

        string? fA1= f.GetFormula(1, 1, "sheet1");
        string? fA2 = f.GetFormula(2, 1, "sheet1");
        string? fB1 = f.GetFormula(1, 2, "sheet1");
        string? fB2 = f.GetFormula(2, 2, "sheet1");

        Assert.AreEqual("SUM(C:D)", fA1);
        Assert.AreEqual("SUM(C:D)", fA2);
        Assert.AreEqual("SUM(D:E)", fB1);
        Assert.AreEqual("SUM(D:E)", fB2);
    }
    [TestMethod]
    public void SharedFormulasShouldNotEffectFullRow()
    {
        ExcelWorksheet.Formulas? f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default) { Index = 0, Formula = "SUM(3:4)", Address = "A1:B2", StartRow = 1, StartCol = 1 };

        string? fA1 = f.GetFormula(1, 1, "sheet1");
        string? fA2 = f.GetFormula(2, 1, "sheet1");
        string? fB1 = f.GetFormula(1, 2, "sheet1");
        string? fB2 = f.GetFormula(2, 2, "sheet1");

        Assert.AreEqual("SUM(3:4)", fA1);
        Assert.AreEqual("SUM(4:5)", fA2);
        Assert.AreEqual("SUM(3:4)", fB1);
        Assert.AreEqual("SUM(4:5)", fB2);
    }
    [TestMethod]
    public void SharedFormulasShouldNotEffectFullSheet()
    {
        ExcelWorksheet.Formulas? f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default) { Index = 0, Formula = "SUM(A:XFD)", Address = "A1:B2", StartRow = 1, StartCol = 1 };

        string? fA1 = f.GetFormula(1, 1, "sheet1");
        string? fA2 = f.GetFormula(2, 1, "sheet1");
        string? fB1 = f.GetFormula(1, 2, "sheet1");
        string? fB2 = f.GetFormula(2, 2, "sheet1");

        Assert.AreEqual("SUM(A:XFD)", fA1);
        Assert.AreEqual("SUM(A:XFD)", fA2);
        Assert.AreEqual("SUM(A:XFD)", fB1);
        Assert.AreEqual("SUM(A:XFD)", fB2);
    }
}