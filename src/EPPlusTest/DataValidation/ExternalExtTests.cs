using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Sparkline;
using System.IO;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing.Slicer;

namespace EPPlusTest.DataValidation;

[TestClass]
public class ExternalExtTests : TestBase
{
    [ClassInitialize]
    public static void Init(TestContext context)
    {
        //_pck = OpenPackage("ExternalReferences.xlsx", true);
        string? outDir = _worksheetPath + "ExternalDataValidations";

        if (!Directory.Exists(outDir))
        {
            _ = Directory.CreateDirectory(outDir);
        }
    }

    //Ensures no save or load errors
    internal static void SaveAndLoadAndSave(in ExcelPackage pck)
    {
        FileInfo? file = pck.File;

        MemoryStream? stream = new MemoryStream();
        pck.SaveAs(stream);

        ExcelPackage? loadedPackage = new ExcelPackage(stream);

        loadedPackage.File = file;

        SaveAndCleanup(loadedPackage);
    }

    internal static void AddDataValidations(ref ExcelWorksheet ws, bool isExtLst = false, string extSheetName = "", bool many = false)
    {
        if (isExtLst)
        {
            IExcelDataValidationInt? intValidation = ws.DataValidations.AddIntegerValidation("A1");
            intValidation.Operator = ExcelDataValidationOperator.equal;
            intValidation.Formula.ExcelFormula = extSheetName + "!A1";
        }
        else
        {
            IExcelDataValidationInt? intValidation = ws.DataValidations.AddIntegerValidation("A1");
            intValidation.Formula.Value = 1;
            intValidation.Formula2.Value = 3;
        }

        if (many)
        {
            IExcelDataValidationTime? timeValidation = ws.DataValidations.AddTimeValidation("B1");
            timeValidation.Operator = ExcelDataValidationOperator.between;

            if (isExtLst)
            {
                timeValidation.Formula.ExcelFormula = extSheetName + "!B1";
                timeValidation.Formula2.ExcelFormula = extSheetName + "!B2";
            }
            else
            {
                timeValidation.Formula.ExcelFormula = "B1";
                timeValidation.Formula.ExcelFormula = "B2";
            }
        }
    }

    [TestMethod]
    public void LocalDataValidationsShouldWorkWithExtLstConditionalFormattings()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\LocalDVExternalCF.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        _ = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

        AddDataValidations(ref ws, false);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void LocalDataValidationsShouldWorkWithManyExtLstConditionalFormattings()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\LocalDVManyExternalCF.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        _ = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);
        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 2, 2, 2), Color.Red);

        AddDataValidations(ref ws, false);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void ManyLocalDataValidationsShouldWorkWithExtLstConditionalFormattings()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\ManyLocalDVExternalCF.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        _ = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);
        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 2, 2, 2), Color.Red);

        AddDataValidations(ref ws, false, "", true);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void ManyLocalDataValidationsShouldWorkWithManyExtLstConditionalFormattings()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\ManyLocalDVManyExternalCF.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        _ = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);
        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

        AddDataValidations(ref ws, false, "", true);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void LocalDataValidationsShouldWorkWithManyExtLstSparklines()
    {
        using ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        _ = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.SparklineGroups.Add(eSparklineType.Line, new ExcelAddress(1, 1, 5, 1), new ExcelAddress(1, 2, 5, 2));
        _ = ws.SparklineGroups.Add(eSparklineType.Line, new ExcelAddress(1, 3, 5, 3), new ExcelAddress(1, 4, 5, 4));

        AddDataValidations(ref ws, false);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void ExtDataValidationsShouldWorkWithExtLstConditionalFormattings()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\ExtDVExternalCF.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        ExcelWorksheet? extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

        AddDataValidations(ref ws, true, extSheet.Name);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void ManyExtDataValidationsShouldWorkWithExtLstConditionalFormattings()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\ManyExtDVExternalCF.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        ExcelWorksheet? extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

        AddDataValidations(ref ws, true, extSheet.Name, true);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void ManyExtDataValidationsShouldWorkWithManyExtLstConditionalFormattings()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\ManyExtDVManyExternalCF.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        ExcelWorksheet? extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);
        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 2, 2, 2), Color.Red);

        AddDataValidations(ref ws, true, extSheet.Name, true);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void ExtDataValidationsShouldWorkWithAllOtherExts()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\ExtDVAllOtherExts.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        ExcelWorksheet? extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["A1:A5"], ws.Cells["B1:B5"]);
        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

        ExcelRange range = ws.Cells[1, 1, 4, 1];
        ExcelTable table = ws.Tables.Add(range, "TestTable");
        table.StyleName = "None";

        _ = ws.Drawings.AddTableSlicer(table.Columns[0]);

        AddDataValidations(ref ws, true, extSheet.Name, true);
        SaveAndLoadAndSave(pck);
    }

    [TestMethod]
    public void LocalDataValidationsShouldWorkWithAllOtherExts()
    {
        using ExcelPackage? pck = OpenPackage("ExternalDataValidations\\LocalDVAllOtherExts.xlsx", true);
        ExcelWorksheet? ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
        _ = pck.Workbook.Worksheets.Add("extAddressSheet");

        _ = ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["A1:A5"], ws.Cells["B1:B5"]);
        _ = ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

        ExcelRange range = ws.Cells[1, 1, 4, 1];
        ExcelTable table = ws.Tables.Add(range, "TestTable");
        table.StyleName = "None";

        _ = ws.Drawings.AddTableSlicer(table.Columns[0]);

        AddDataValidations(ref ws, false);
        SaveAndLoadAndSave(pck);
    }
}