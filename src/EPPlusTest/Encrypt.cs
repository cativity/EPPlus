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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace EPPlusTest;

[TestClass]
public class Encrypt : TestBase
{
    //[TestMethod]
    //[Ignore]
    //public void ReadWriteEncrypt()
    //{
    //    using (ExcelPackage pck = new ExcelPackage(new FileInfo(@"Test\Drawing.xlsx"), true))
    //    {
    //        pck.Encryption.Password = "EPPlus";
    //        pck.Encryption.Algorithm = EncryptionAlgorithm.AES192;
    //        pck.Workbook.Protection.SetPassword("test");
    //        pck.Workbook.Protection.LockStructure = true;
    //        pck.Workbook.Protection.LockWindows = true;

    //        pck.SaveAs(new FileInfo(@"Test\DrawingEncr.xlsx"));
    //    }

    //    using (ExcelPackage pck = new ExcelPackage(new FileInfo(_worksheetPath + @"\DrawingEncr.xlsx"), true, "EPPlus"))
    //    {
    //        pck.Encryption.IsEncrypted = false;
    //        pck.SaveAs(new FileInfo(_worksheetPath + @"\DrawingNotEncr.xlsx"));
    //    }

    //    FileStream fs = new FileStream(_worksheetPath + @"\DrawingEncr.xlsx", FileMode.Open, FileAccess.ReadWrite);

    //    using (ExcelPackage pck = new ExcelPackage(fs, "EPPlus"))
    //    {
    //        pck.Encryption.IsEncrypted = false;
    //        pck.SaveAs(new FileInfo(_worksheetPath + @"DrawingNotEncr.xlsx"));
    //    }
    //}

    //[TestMethod]
    //[Ignore]
    //public void WriteEncrypt()
    //{
    //    ExcelPackage package = new ExcelPackage();

    //    //Load the sheet with one string column, one date column and a few random numbers.
    //    ExcelWorksheet? ws = package.Workbook.Worksheets.Add("First line test");

    //    ws.Cells[1, 1].Value = "1; 1";
    //    ws.Cells[2, 1].Value = "2; 1";
    //    ws.Cells[1, 2].Value = "1; 2";
    //    ws.Cells[2, 2].Value = "2; 2";

    //    ws.Row(1).Style.Font.Bold = true;
    //    ws.Column(1).Style.Font.Bold = true;

    //    //package.Encryption.Algorithm = EncryptionAlgorithm.AES256;
    //    //package.SaveAs(new FileInfo(@"c:\temp\encrTest.xlsx"), "ABxsw23edc");
    //    package.Encryption.Password = "test";
    //    package.Encryption.IsEncrypted = true;
    //    package.SaveAs(new FileInfo(@"c:\temp\encrTest.xlsx"));
    //}

    //[TestMethod]
    //[Ignore]
    //public void WriteProtect()
    //{
    //    ExcelPackage package = new ExcelPackage(new FileInfo(@"c:\temp\workbookprot2.xlsx"), "");

    //    //Load the sheet with one string column, one date column and a few random numbers.
    //    //package.Workbook.Protection.LockWindows = true;
    //    //package.Encryption.IsEncrypted = true;
    //    package.Workbook.Protection.SetPassword("t");
    //    package.Workbook.Protection.LockStructure = true;
    //    package.Workbook.View.Left = 585;
    //    package.Workbook.View.Top = 150;

    //    package.Workbook.View.Width = 17310;
    //    package.Workbook.View.Height = 38055;
    //    ExcelWorksheet? ws = package.Workbook.Worksheets.Add("First line test");

    //    ws.Cells[1, 1].Value = "1; 1";
    //    ws.Cells[2, 1].Value = "2; 1";
    //    ws.Cells[1, 2].Value = "1; 2";
    //    ws.Cells[2, 2].Value = "2; 2";

    //    package.SaveAs(new FileInfo(@"c:\temp\workbookprot2.xlsx"));
    //}

    //[TestMethod]
    //[Ignore]
    //public void DecrypTest()
    //{
    //    ExcelPackage? p = new ExcelPackage(new FileInfo(@"c:\temp\encr.xlsx"), "test");

    //    string? n = p.Workbook.Worksheets[1].Name;
    //    p.Encryption.Password = null;
    //    p.SaveAs(new FileInfo(@"c:\temp\encrNew.xlsx"));
    //}

    //[TestMethod, Ignore]
    //public void DecrypTestBug()
    //{
    //    ExcelPackage? p = new ExcelPackage(new FileInfo(@"c:\temp\bug\TestExcel_2040.xlsx"), "");

    //    string? n = p.Workbook.Worksheets[1].Name;
    //    p.Encryption.Password = null;
    //    p.SaveAs(new FileInfo(@"c:\temp\encrNew.xlsx"));
    //}

    //[TestMethod]
    //[Ignore]
    //public void EncrypTest()
    //{
    //    FileInfo? f = new FileInfo(@"c:\temp\encrwrite.xlsx");

    //    if (f.Exists)
    //    {
    //        f.Delete();
    //    }

    //    ExcelPackage? p = new ExcelPackage(f);

    //    p.Workbook.Protection.SetPassword("");
    //    p.Workbook.Protection.LockStructure = true;
    //    p.Encryption.Version = EncryptionVersion.Agile;

    //    ExcelWorksheet? ws = p.Workbook.Worksheets.Add("Sheet1");

    //    for (int r = 1; r < 1000; r++)
    //    {
    //        ws.Cells[r, 1].Value = r;
    //    }

    //    p.Save();
    //}

    [TestMethod]
    public void ValidateStaticEnryptionMethods()
    {
        using ExcelPackage? p = new ExcelPackage();
        _ = p.Workbook.Worksheets.Add("Sheet1");
        p.Save();

        MemoryStream? ep = ExcelEncryption.EncryptPackage(p.Stream, "EPPlus");
        MemoryStream? dp = ExcelEncryption.DecryptPackage(ep, "EPPlus");

        using ExcelPackage? p2 = new ExcelPackage(dp);
        Assert.AreEqual(p.Workbook.Worksheets.Count, p2.Workbook.Worksheets.Count);
        Assert.AreEqual(p.Workbook.Worksheets[0].Name, p2.Workbook.Worksheets[0].Name);
    }
}