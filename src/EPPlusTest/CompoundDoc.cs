/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/

using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Linq;
using OfficeOpenXml.VBA;

namespace EPPlusTest;

/// <summary>
/// Summary description for CompoundDoc
/// </summary>
[TestClass]
public class CompoundDoc
{
    public CompoundDoc()
    {
        //
        // TODO: Add constructor logic here
        //
    }

    private TestContext testContextInstance;

    /// <summary>
    ///Gets or sets the test context which provides
    ///information about and functionality for the current test run.
    ///</summary>
    public TestContext TestContext
    {
        get => this.testContextInstance;
        set => this.testContextInstance = value;
    }

    #region Additional test attributes

    //
    // You can use the following additional attributes as you write your tests:
    //
    // Use ClassInitialize to run code before running the first test in the class
    // [ClassInitialize()]
    // public static void MyClassInitialize(TestContext testContext) { }
    //
    // Use ClassCleanup to run code after all tests in a class have run
    // [ClassCleanup()]
    // public static void MyClassCleanup() { }
    //
    // Use TestInitialize to run code before running each test 
    // [TestInitialize()]
    // public void MyTestInitialize() { }
    //
    // Use TestCleanup to run code after each test has run
    // [TestCleanup()]
    // public void MyTestCleanup() { }
    //

    #endregion

    //[TestMethod, Ignore]
    //public void Read()
    //{
    //    //var doc = File.ReadAllBytes(@"c:\temp\vbaProject.bin");
    //    byte[]? doc = File.ReadAllBytes(@"c:\temp\vba.bin");
    //    CompoundDocumentFile? cd = new CompoundDocumentFile(doc);
    //    MemoryStream? ms = new MemoryStream();
    //    cd.Write(ms);
    //    printitems(cd.RootItem);
    //    File.WriteAllBytes(@"c:\temp\vba.bin", ms.ToArray());
    //}

    private static void printitems(CompoundDocumentItem item)
    {
        File.AppendAllText(@"c:\temp\items.txt", item.Name + "\t");

        foreach (CompoundDocumentItem? c in item.Children)
        {
            printitems(c);
        }
    }

    //[TestMethod, Ignore]
    //public void WriteReadCompundDoc()
    //{
    //    for (int i = 1; i < 50; i++)
    //    {
    //        byte[]? b = CreateFile(i);
    //        ReadFile(b, i);
    //        GC.Collect();
    //    }

    //    for (int i = 5; i < 20; i++)
    //    {
    //        byte[]? b = CreateFile(i * 50);
    //        ReadFile(b, i * 50);
    //        GC.Collect();
    //    }
    //}

    private static void ReadFile(byte[] b, int noSheets)
    {
        MemoryStream? ms = new MemoryStream(b);
        using ExcelPackage? p = new ExcelPackage(ms);
        Assert.AreEqual(p.Workbook.VbaProject.Modules.Count, noSheets + 2);
        Assert.AreEqual(noSheets, p.Workbook.Worksheets.Count);
    }

    public static byte[] CreateFile(int noSheets)
    {
        using ExcelPackage? package = new ExcelPackage();

        IEnumerable<string>? sheets = Enumerable.Range(1, noSheets) //460
                                                .Select(x => $"Sheet{x}");

        foreach (string? sheet in sheets)
        {
            _ = package.Workbook.Worksheets.Add(sheet);
        }

        package.Workbook.CreateVBAProject();
        package.Workbook.VbaProject.Modules.AddModule("Module1").Code = "\r\nPublic Sub SayHello()\r\nMsgBox(\"Hello\")\r\nEnd Sub\r\n";

        return package.GetAsByteArray();
    }

    //[TestMethod, Ignore]
    //public void ReadEncLong()
    //{
    //    byte[]? doc = File.ReadAllBytes(@"c:\temp\EncrDocRead.xlsx");
    //    CompoundDocumentFile? cd = new CompoundDocumentFile(doc);
    //    MemoryStream? ms = new MemoryStream();
    //    cd.Write(ms);

    //    File.WriteAllBytes(@"c:\temp\vba.xlsx", ms.ToArray());
    //}

    //[TestMethod, Ignore]
    //public void ReadVba()
    //{
    //    ExcelPackage? p = new ExcelPackage(new FileInfo(@"c:\temp\pricecheck.xlsm"));
    //    ExcelVbaProject? vba = p.Workbook.VbaProject;
    //    p.SaveAs(new FileInfo(@"c:\temp\pricecheckSaved.xlsm"));
    //}

    static FileInfo TempFile(string name)
    {
        string? baseFolder = Path.Combine(@"c:\temp\bug\");

        return new FileInfo(Path.Combine(baseFolder, name));
    }

    //[TestMethod, Ignore]
    //public void Issue131()
    //{
    //    FileInfo? src = TempFile("report.xlsm");

    //    if (src.Exists)
    //    {
    //        src.Delete();
    //    }

    //    ExcelPackage? package = new ExcelPackage(src);

    //    IEnumerable<string>? sheets = Enumerable.Range(1, 500) //460
    //                                            .Select(x => $"Sheet{x}");

    //    foreach (string? sheet in sheets)
    //    {
    //        _ = package.Workbook.Worksheets.Add(sheet);
    //    }

    //    package.Workbook.CreateVBAProject();
    //    package.Workbook.VbaProject.Modules.AddModule("Module1").Code = "\r\nPublic Sub SayHello()\r\nMsgBox(\"Hello\")\r\nEnd Sub\r\n";

    //    package.Save();
    //}

    //[TestMethod, Ignore]
    //public void Sample7EncrLargeTest()
    //{
    //    int Rows = 1000000;
    //    int colMult = 20;
    //    FileInfo newFile = new FileInfo(@"C:\temp\bug\sample7compdoctest.xlsx");

    //    if (newFile.Exists)
    //    {
    //        newFile.Delete(); // ensures we create a new workbook
    //        newFile = new FileInfo(@"C:\temp\bug\sample7compdoctest.xlsx");
    //    }

    //    using (ExcelPackage package = new ExcelPackage())
    //    {
    //        Console.WriteLine("{0:HH.mm.ss}\tStarting...", DateTime.Now);

    //        //Load the sheet with one string column, one date column and a few random numbers.
    //        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("Performance Test");

    //        //Format all cells
    //        ExcelRange cols = ws.Cells["A:XFD"];
    //        cols.Style.Fill.PatternType = ExcelFillStyle.Solid;
    //        cols.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

    //        Random? rnd = new Random();

    //        for (int row = 1; row <= Rows; row++)
    //        {
    //            for (int c = 0; c < colMult * 5; c += 5)
    //            {
    //                ws.SetValue(row, 1 + c, row); //The SetValue method is a little bit faster than using the Value property
    //                ws.SetValue(row, 2 + c, string.Format("Row {0}", row));
    //                ws.SetValue(row, 3 + c, DateTime.Today.AddDays(row));
    //                ws.SetValue(row, 4 + c, rnd.NextDouble() * 10000);
    //            }
    //        }

    //        int endC = colMult * 5;
    //        ws.Cells[1, endC, Rows, endC].FormulaR1C1 = "RC[-4]+RC[-1]";

    //        //Add a sum at the end
    //        ws.Cells[Rows + 1, endC].Formula = string.Format("Sum({0})", new ExcelAddress(1, 5, Rows, 5).Address);
    //        ws.Cells[Rows + 1, endC].Style.Font.Bold = true;
    //        ws.Cells[Rows + 1, endC].Style.Numberformat.Format = "#,##0.00";

    //        Console.WriteLine("{0:HH.mm.ss}\tWriting row {1}...", DateTime.Now, Rows);
    //        Console.WriteLine("{0:HH.mm.ss}\tFormatting...", DateTime.Now);

    //        //Format the date and numeric columns
    //        ws.Cells[1, 1, Rows, 1].Style.Numberformat.Format = "#,##0";
    //        ws.Cells[1, 3, Rows, 3].Style.Numberformat.Format = "YYYY-MM-DD";
    //        ws.Cells[1, 4, Rows, 5].Style.Numberformat.Format = "#,##0.00";

    //        Console.WriteLine("{0:HH.mm.ss}\tInsert a row at the top...", DateTime.Now);

    //        //Insert a row at the top. Note that the formula-addresses are shifted down
    //        ws.InsertRow(1, 1);

    //        //Write the headers and style them
    //        ws.Cells["A1"].Value = "Index";
    //        ws.Cells["B1"].Value = "Text";
    //        ws.Cells["C1"].Value = "Date";
    //        ws.Cells["D1"].Value = "Number";
    //        ws.Cells["E1"].Value = "Formula";
    //        ws.View.FreezePanes(2, 1);

    //        using (ExcelRange? rng = ws.Cells["A1:E1"])
    //        {
    //            rng.Style.Font.Bold = true;
    //            rng.Style.Font.Color.SetColor(Color.White);
    //            rng.Style.WrapText = true;
    //            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    //            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    //            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
    //            rng.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
    //        }

    //        //Calculate (Commented away thisk, it was a bit time consuming... /MA)
    //        // Console.WriteLine("{0:HH.mm.ss}\tCalculate formulas...", DateTime.Now);
    //        // ws.Calculate();

    //        Console.WriteLine("{0:HH.mm.ss}\tAutofit columns and lock and format cells...", DateTime.Now);
    //        ws.Cells[Rows - 100, 1, Rows, 5].AutoFitColumns(5); //Auto fit using the last 100 rows with minimum width 5

    //        ws.Column(5).Width =
    //            15; //We need to set the width for column F manually since the end sum formula is the widest cell in the column (EPPlus don't calculate any forumlas, so no output text is avalible). 

    //        //Now we set the sheetprotection and a password.
    //        ws.Cells[2, 3, Rows + 1, 4].Style.Locked = false;
    //        ws.Cells[2, 3, Rows + 1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
    //        ws.Cells[2, 3, Rows + 1, 4].Style.Fill.BackgroundColor.SetColor(Color.White);
    //        ws.Cells[1, 5, Rows + 2, 5].Style.Hidden = true; //Hide the formula

    //        ws.Protection.SetPassword("EPPlus");

    //        ws.Select("C2");
    //        Console.WriteLine("{0:HH.mm.ss}\tSaving...", DateTime.Now);
    //        package.Compression = CompressionLevel.BestSpeed;
    //        package.Encryption.IsEncrypted = true;
    //        package.SaveAs(newFile);
    //    }

    //    Console.WriteLine("{0:HH.mm.ss}\tDone!!", DateTime.Now);
    //}

    //[TestMethod, Ignore]
    //public void ReadPerfTest()
    //{
    //    ExcelPackage? p = new ExcelPackage(new FileInfo(@"c:\temp\bug\sample7compdoctest.xlsx"), "");

    //    //var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\sample7compdoctest_4.5.xlsx"), "");
    //    //var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\sample7compdoctest.310k.xlsx"), "");
    //}

    //[TestMethod, Ignore]
    //public void ReadVbaIssue107()
    //{
    //    //var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\report.xlsm"));
    //    //var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\report411.xlsm"));
    //    ExcelPackage? p = new ExcelPackage(new FileInfo(@"c:\temp\bug\sample7.xlsx"), "");
    //    ExcelVbaProject? vba = p.Workbook.VbaProject;
    //}
}