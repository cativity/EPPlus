using EPPlusTest;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace EPPlusTest.Core;

[TestClass]
public class ExternalLinksTest : TestBase
{
    //static ExcelPackage _pck;
    [ClassInitialize]
    public static void Init(TestContext context)
    {
        //_pck = OpenPackage("ExternalReferences.xlsx", true);
        string? outDir = _worksheetPath + "ExternalReferences";

        if (!Directory.Exists(outDir))
        {
            Directory.CreateDirectory(outDir);
        }

        foreach (string? f in Directory.GetFiles(_testInputPath + "ExternalReferences"))
        {
            File.Copy(f, outDir + "\\" + new FileInfo(f).Name, true);
        }
    }

    [ClassCleanup]
    public static void Cleanup()
    {
        //var dirName = _pck.File.DirectoryName;
        //var fileName = _pck.File.FullName;

        //SaveAndCleanup(_pck);

        //if (File.Exists(fileName)) File.Copy(fileName, dirName + "\\ExternalReferencesRead.xlsx", true);
    }

    [TestMethod]
    public void OpenAndReadExternalLink()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferences\\ExtRef.xlsx");

        Assert.AreEqual(2, p.Workbook.ExternalLinks.Count);

        Assert.AreEqual(1D, p.Workbook.ExternalLinks[0].As.ExternalWorkbook.CachedWorksheets["sheet1"].CellValues["A2"].Value);
        Assert.AreEqual(12D, p.Workbook.ExternalLinks[0].As.ExternalWorkbook.CachedWorksheets["sheet1"].CellValues["C3"].Value);

        int c = 0;

        foreach (ExcelExternalCellValue? cell in p.Workbook.ExternalLinks[0].As.ExternalWorkbook.CachedWorksheets["sheet1"].CellValues)
        {
            Assert.IsNotNull(cell.Value);
            c++;
        }

        Assert.AreEqual(11, c);
    }

    [TestMethod]
    public void OpenAndCalculateExternalLinkFromCache()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferences\\ExtRef.xlsx");

        p.Workbook.ClearFormulaValues();
        p.Workbook.Calculate();

        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        Assert.AreEqual(2D, ws.Cells["E2"].Value);
        Assert.AreEqual(4D, ws.Cells["F2"].Value);
        Assert.AreEqual(6D, ws.Cells["G2"].Value);

        Assert.AreEqual(8D, ws.Cells["E3"].Value);
        Assert.AreEqual(16D, ws.Cells["F3"].Value);
        Assert.AreEqual(24D, ws.Cells["G3"].Value);

        Assert.AreEqual(20D, ws.Cells["H5"].Value);
        Assert.AreEqual(117D, ws.Cells["K5"].Value);

        Assert.AreEqual(111D, ws.Cells["H8"].Value);
        Assert.IsInstanceOfType(ws.Cells["J8"].Value, typeof(ExcelErrorValue));
        Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)ws.Cells["J8"].Value).Type);

        Assert.AreEqual(3D, ws.Cells["E10"].Value);
        Assert.IsInstanceOfType(ws.Cells["F10"].Value, typeof(ExcelErrorValue));
        Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)ws.Cells["F10"].Value).Type);
    }

    [TestMethod]
    public void OpenAndCalculateExternalLinkFromPackage()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferences\\ExtRef.xlsx");

        p.Workbook.ExternalLinks.Directories.Add(new DirectoryInfo(_testInputPathOptional));
        p.Workbook.ExternalLinks.LoadWorkbooks();
        p.Workbook.ExternalLinks[0].As.ExternalWorkbook.Package.Workbook.Calculate();
        p.Workbook.ClearFormulaValues();
        p.Workbook.Calculate();

        ExcelWorksheet? ws = p.Workbook.Worksheets[0];
        Assert.AreEqual(3D, ws.Cells["D1"].Value);
        Assert.AreEqual(2D, ws.Cells["E2"].Value);
        Assert.AreEqual(2D, ws.Cells["E2"].Value);
        Assert.AreEqual(4D, ws.Cells["F2"].Value);
        Assert.AreEqual(6D, ws.Cells["G2"].Value);

        Assert.AreEqual(8D, ws.Cells["E3"].Value);
        Assert.AreEqual(16D, ws.Cells["F3"].Value);
        Assert.AreEqual(24D, ws.Cells["G3"].Value);

        Assert.AreEqual(117D, ws.Cells["K5"].Value);

        Assert.AreEqual(111D, ws.Cells["H8"].Value);
        Assert.AreEqual(20D, ws.Cells["J8"].Value);

        Assert.AreEqual(3D, ws.Cells["E10"].Value);
        Assert.AreEqual(19D, ws.Cells["F10"].Value);
    }

    [TestMethod]
    public void DeleteExternalLink()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferences\\ExtRef.xlsx");

        Assert.AreEqual(2, p.Workbook.ExternalLinks.Count);

        p.Workbook.ExternalLinks.RemoveAt(0);

        SaveWorkbook("ExtRefDeleted.xlsx", p);
    }

    [TestMethod]
    public void OpenAndReadExternalReferences1()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

        Assert.AreEqual(62, p.Workbook.ExternalLinks.Count);

        int c = 0;

        foreach (ExcelExternalCellValue? cell in p.Workbook.ExternalLinks[0].As.ExternalWorkbook.CachedWorksheets[0].CellValues)
        {
            Assert.IsNotNull(cell.Value);
            c++;
        }

        Assert.AreEqual(104, c);
    }

    [TestMethod]
    public void DeleteExternalLinks1()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

        p.Workbook.ExternalLinks.RemoveAt(0);
        p.Workbook.ExternalLinks.RemoveAt(8);
        p.Workbook.ExternalLinks.RemoveAt(5);

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void OpenAndReadExternalLinks2()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

        Assert.AreEqual(204, p.Workbook.ExternalLinks.Count);

        int c = 0;

        foreach (ExcelExternalCellValue? cell in p.Workbook.ExternalLinks[0].As.ExternalWorkbook.CachedWorksheets[0].CellValues)
        {
            Assert.IsNotNull(cell.Value);
            c++;
        }

        Assert.AreEqual(104, c);
    }

    [TestMethod]
    public void OpenAndDeleteExternalLinks2()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

        Assert.AreEqual(204, p.Workbook.ExternalLinks.Count);
        p.Workbook.ExternalLinks.RemoveAt(103);
        Assert.AreEqual(203, p.Workbook.ExternalLinks.Count);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void OpenAndCalculateExternalLinks1()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

        p.Workbook.Calculate();
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void OpenAndCalculateExternalLinks2()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

        Assert.AreEqual(204, p.Workbook.ExternalLinks.Count);
        p.Workbook.Calculate();
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void OpenAndClearExternalLinks1()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

        Assert.AreEqual(62, p.Workbook.ExternalLinks.Count);
        p.Workbook.ExternalLinks.Clear();
        Assert.AreEqual(0, p.Workbook.ExternalLinks.Count);
        SaveWorkbook("ExternalReferencesText1_Cleared.xlsx", p);
    }

    [TestMethod]
    public void OpenAndClearExternalLinks2()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

        Assert.AreEqual(204, p.Workbook.ExternalLinks.Count);
        p.Workbook.ExternalLinks.Clear();
        Assert.AreEqual(0, p.Workbook.ExternalLinks.Count);
        SaveWorkbook("ExternalReferencesText2_Cleared.xlsx", p);
    }

    [TestMethod]
    public void OpenAndClearExternalLinks3()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText3.xlsx");

        Assert.AreEqual(63, p.Workbook.ExternalLinks.Count);
        p.Workbook.ExternalLinks.Clear();
        Assert.AreEqual(0, p.Workbook.ExternalLinks.Count);
        SaveWorkbook("ExternalReferencesText3_Cleared.xlsx", p);
    }

    [TestMethod]
    public void OpenAndReadExternalLinks3()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText3.xlsx");

        Assert.AreEqual(63, p.Workbook.ExternalLinks.Count);

        int c = 0;

        foreach (ExcelExternalCellValue? cell in p.Workbook.ExternalLinks[0].As.ExternalWorkbook.CachedWorksheets[0].CellValues)
        {
            Assert.IsNotNull(cell.Value);
            c++;
        }

        Assert.AreEqual(104, c);
    }

    [TestMethod]
    public void OpenAndCalculateExternalLink3()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferencesText3.xlsx");

        p.Workbook.Calculate();
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void OpenAndReadExternalLinkDdeOle()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferences\\dde.xlsx");

        Assert.AreEqual(6, p.Workbook.ExternalLinks.Count);

        Assert.AreEqual(eExternalLinkType.DdeLink, p.Workbook.ExternalLinks[0].ExternalLinkType);
        p.Workbook.ExternalLinks.LoadWorkbooks();

        ExcelExternalWorkbook? book3 = p.Workbook.ExternalLinks[3].As.ExternalWorkbook;
        Assert.AreEqual(p.File.DirectoryName + "\\fromwb1.xlsx", book3.File.FullName, true);
        Assert.IsNotNull(book3.Package);
        ExcelExternalWorkbook? book4 = p.Workbook.ExternalLinks[4].As.ExternalWorkbook;
        Assert.AreEqual(p.File.DirectoryName + "\\extref.xlsx", book4.File.FullName, true);
        Assert.IsNotNull(book4.Package);
        SaveWorkbook("dde.xlsx", p);
    }

    [TestMethod]
    public void UpdateCacheShouldBeSameAsExcel()
    {
        ExcelPackage? p = OpenTemplatePackage("ExternalReferences\\ExtRef.xlsx");

        ExcelExternalWorkbook? er = p.Workbook.ExternalLinks[0].As.ExternalWorkbook;
        Dictionary<string, object>? excelCache = GetExternalCache(er);

        p.Workbook.ExternalLinks[0].As.ExternalWorkbook.UpdateCache();
        Dictionary<string, object>? epplusCache = GetExternalCache(er);

        foreach (string? key in excelCache.Keys)
        {
            if (epplusCache.ContainsKey(key))
            {
                //We remove any single quotes as excel adds seems to add ' to all worksheet names.
                //We also remove any prefixing equal sign in teh GetExternalCache method.
                Assert.AreEqual(excelCache[key].ToString().Replace("\'", "").ToString(), epplusCache[key].ToString().Replace("\'", ""));
            }
            else
            {
                Assert.Fail($"Key:{key} are missing in EPPlus cache.");
            }
        }

        foreach (string? key in epplusCache.Keys)
        {
            if (!excelCache.ContainsKey(key))
            {
                Assert.Fail($"Key:{key} are missing in EPPlus cache.");
            }
        }

        SaveWorkbook("ExtRef_Updated.xlsx", p);
    }

    [TestMethod]
    public void AddExternalLinkShouldBeSameAsExcel()
    {
        ExcelPackage? p = OpenPackage("AddedExtRef.xlsx", true);
        ExcelWorksheet? ws1 = CreateWorksheet1(p);
        ExcelWorksheet? ws2 = p.Workbook.Worksheets.Add("Sheet2");

        ws2.Cells["A1"].Value = 3;
        ws2.Names.Add("SheetDefinedName", ws2.Cells["A1"]);

        ws1.Cells["D2"].Formula = "Sheet2!SheetDefinedName";
        ws1.Cells["E2"].Formula = "Table1[[#This Row],[a]]+[1]Sheet1!$A2";
        ws1.Cells["F2"].Formula = "Table1[[#This Row],[b]]+[1]Sheet1!$B2";
        ws1.Cells["G2"].Formula = "Table1[[#This Row],[c]]+[1]Sheet1!$C2";
        ws1.Cells["E3"].Formula = "Table1[[#This Row],[a]]+[1]Sheet1!$A3";
        ws1.Cells["F3"].Formula = "Table1[[#This Row],[b]]+[1]Sheet1!$B3";
        ws1.Cells["G3"].Formula = "Table1[[#This Row],[c]]+'[1]Sheet1'!$C3";
        ws1.Cells["G4"].Formula = "Table1[[#This Row],[c]]+'[1]Sheet8888'!$C3";
        ExcelExternalWorkbook? er = p.Workbook.ExternalLinks.AddExternalWorkbook(new FileInfo(_testInputPath + "externalreferences\\FromWB1.xlsx"));

        ws1.Cells["G5"].Formula = $"[{er.Index}]Sheet1!FromF2*[{er.Index}]!CellH5";

        er.UpdateCache();
        ws1.Calculate();
        p.Workbook.ExternalLinks.UpdateCaches();

        Assert.IsInstanceOfType(ws1.Cells["G4"].Value, typeof(ExcelErrorValue));
        Assert.AreEqual(2220D, ws1.Cells["G5"].Value);
        SaveAndCleanup(p);
    }

    [TestMethod]
    public void AddExternalWorkbookNoUpdate()
    {
        ExcelPackage? p = OpenPackage("AddedExtRefNoUpdate.xlsx", true);
        ExcelWorksheet? ws1 = CreateWorksheet1(p);
        ExcelWorksheet? ws2 = p.Workbook.Worksheets.Add("Sheet2");

        ws2.Cells["A1"].Value = 3;
        ws2.Names.Add("SheetDefinedName", ws2.Cells["A1"]);

        ws1.Cells["D2"].Formula = "Sheet2!SheetDefinedName";
        ws1.Cells["E2"].Formula = "Table1[[#This Row],[a]]+[1]Sheet1!$A2";
        ws1.Cells["F2"].Formula = "Table1[[#This Row],[b]]+[1]Sheet1!$B2";
        ws1.Cells["G2"].Formula = "Table1[[#This Row],[c]]+[1]Sheet1!$C2";
        ws1.Cells["E3"].Formula = "Table1[[#This Row],[a]]+[1]Sheet1!$A3";
        ws1.Cells["F3"].Formula = "Table1[[#This Row],[b]]+[1]Sheet1!$B3";
        ws1.Cells["G3"].Formula = "Table1[[#This Row],[c]]+'[1]Sheet1'!$C3";
        ExcelExternalWorkbook? er = p.Workbook.ExternalLinks.AddExternalWorkbook(new FileInfo(_testInputPath + "externalreferences\\FromWB1.xlsx"));

        ws1.Cells["G5"].Formula = $"[{er.Index}]Sheet1!FromF2*[{er.Index}]!CellH5";
        ws1.Cells["G6"].Formula = $"'[FromWB1.xlsx]Sheet1'!FromF2*[FromWB1.xlsx]Sheet1!H6";
        ws1.Cells["G7"].Formula = $"'[c:/epplusTest/Testoutput/externalreferences/FromWB1.xlsx]Sheet1'!FromF2*[FromWB1.xlsx]Sheet1!H6";

        SaveAndCleanup(p);
    }

    [TestMethod]
    public void AddExternalWorkbookWithChartCache()
    {
        ExcelPackage? p = OpenPackage("AddedExtRefChart.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("SheetWithChart");

        ExcelExternalWorkbook? er = p.Workbook.ExternalLinks.AddExternalWorkbook(new FileInfo(_testInputPath + "externalreferences\\FromWB1.xlsx"));
        ExcelLineChart? chart = ws.Drawings.AddLineChart("line1", eLineChartType.Line);
        ExcelLineChartSerie? serie = chart.Series.Add("[1]Sheet1!A2:A3", "[1]Sheet1!B2:B3");
        er.UpdateCache();
        serie.CreateCache();

        SaveAndCleanup(p);
    }

    private static ExcelWorksheet CreateWorksheet1(ExcelPackage p)
    {
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");
        ws.SetValue(1, 1, "a");
        ws.SetValue(1, 2, "b");
        ws.SetValue(1, 3, "c");
        ws.SetValue(2, 1, 1D);
        ws.SetValue(2, 2, 2D);
        ws.SetValue(2, 3, 3D);
        ws.SetValue(3, 1, 4D);
        ws.SetValue(3, 2, 8D);
        ws.SetValue(3, 3, 12D);

        ExcelTable? tbl = ws.Tables.Add(ws.Cells["A1:C3"], "Table1");

        //Create Table
        tbl.TableStyle = TableStyles.Medium2;

        return ws;
    }

    private static Dictionary<string, object> GetExternalCache(ExcelExternalWorkbook ewb)
    {
        Dictionary<string, object>? d = new Dictionary<string, object>();

        foreach (ExcelExternalWorksheet ws in ewb.CachedWorksheets)
        {
            foreach (ExcelExternalCellValue v in ws.CellValues)
            {
                d.Add(ws.Name + v.Address, v.Value);
            }

            foreach (ExcelExternalDefinedName n in ws.CachedNames)
            {
                if (n.RefersTo.StartsWith("="))
                {
                    d.Add(ws.Name + n.Name, n.RefersTo.Substring(1));
                }
                else
                {
                    d.Add(ws.Name + n.Name, n.RefersTo);
                }
            }
        }

        foreach (ExcelExternalDefinedName n in ewb.CachedNames)
        {
            if (n.RefersTo.StartsWith("="))
            {
                d.Add(n.Name, n.RefersTo.Substring(1));
            }
            else
            {
                d.Add(n.Name, n.RefersTo);
            }
        }

        return d;
    }

    [TestMethod]
    public void RichTextClear()
    {
        using ExcelPackage? p = OpenPackage("RichText.xlsx", true);
        ExcelWorksheet? ws = p.Workbook.Worksheets.Add("SheetWithChart");
        ws.Cells["A1:A5"].FillNumber(1, 1);
        ws.Cells["A1:A5"].Style.Font.Bold = true;
        ws.Cells["A1:A5"].RichText.Clear();
        ws.Cells["A1:A5"].FillNumber(1, 1);
        SaveAndCleanup(p);
    }
}