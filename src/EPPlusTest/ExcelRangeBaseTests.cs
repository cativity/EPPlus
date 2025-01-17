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

namespace EPPlusTest;

[TestClass]
public class ExcelRangeBaseTests : TestBase
{
    [TestMethod]
    public void CopyCopiesCommentsFromSingleCellRanges()
    {
        InitBase();
        ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
        ExcelRange? sourceExcelRange = ws1.Cells[3, 3];
        Assert.IsNull(sourceExcelRange.Comment);
        _ = sourceExcelRange.AddComment("Testing comment 1", "test1");
        Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
        Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
        ExcelRange? destinationExcelRange = ws1.Cells[5, 5];
        Assert.IsNull(destinationExcelRange.Comment);
        sourceExcelRange.Copy(destinationExcelRange);

        // Assert the original comment is intact.
        Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
        Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);

        // Assert the comment was copied.
        Assert.AreEqual("test1", destinationExcelRange.Comment.Author);
        Assert.AreEqual("Testing comment 1", destinationExcelRange.Comment.Text);
    }

    [TestMethod]
    public void CopyCopiesCommentsFromMultiCellRanges()
    {
        InitBase();
        ExcelPackage? pck = new ExcelPackage();
        ExcelWorksheet? ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
        ExcelRange? sourceExcelRangeC3 = ws1.Cells[3, 3];
        ExcelRange? sourceExcelRangeD3 = ws1.Cells[3, 4];
        ExcelRange? sourceExcelRangeE3 = ws1.Cells[3, 5];
        Assert.IsNull(sourceExcelRangeC3.Comment);
        Assert.IsNull(sourceExcelRangeD3.Comment);
        Assert.IsNull(sourceExcelRangeE3.Comment);
        _ = sourceExcelRangeC3.AddComment("Testing comment 1", "test1");
        _ = sourceExcelRangeD3.AddComment("Testing comment 2", "test1");
        _ = sourceExcelRangeE3.AddComment("Testing comment 3", "test1");
        Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
        Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
        Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
        Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
        Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
        Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);

        // Copy the full row to capture each cell at once.
        Assert.IsNull(ws1.Cells[5, 3].Comment);
        Assert.IsNull(ws1.Cells[5, 4].Comment);
        Assert.IsNull(ws1.Cells[5, 5].Comment);
        ws1.Cells["3:3"].Copy(ws1.Cells["5:5"]);

        // Assert the original comments are intact.
        Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
        Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
        Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
        Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
        Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
        Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);

        // Assert the comments were copied.
        ExcelRange? destinationExcelRangeC5 = ws1.Cells[5, 3];
        ExcelRange? destinationExcelRangeD5 = ws1.Cells[5, 4];
        ExcelRange? destinationExcelRangeE5 = ws1.Cells[5, 5];
        Assert.AreEqual("test1", destinationExcelRangeC5.Comment.Author);
        Assert.AreEqual("Testing comment 1", destinationExcelRangeC5.Comment.Text);
        Assert.AreEqual("test1", destinationExcelRangeD5.Comment.Author);
        Assert.AreEqual("Testing comment 2", destinationExcelRangeD5.Comment.Text);
        Assert.AreEqual("test1", destinationExcelRangeE5.Comment.Author);
        Assert.AreEqual("Testing comment 3", destinationExcelRangeE5.Comment.Text);
    }

    [TestMethod]
    public void SettingAddressHandlesMultiAddresses()
    {
        using ExcelPackage package = new ExcelPackage();
        ExcelWorksheet? worksheet = package.Workbook.Worksheets.Add("Sheet1");
        ExcelNamedRange? name = package.Workbook.Names.Add("Test", worksheet.Cells[3, 3]);
        name.Address = "Sheet1!C3";
        name.Address = "Sheet1!D3";
        Assert.IsNull(name.Addresses);
        name.Address = "C3:D3,E3:F3";
        Assert.IsNotNull(name.Addresses);
        name.Address = "Sheet1!C3";
        Assert.IsNull(name.Addresses);
    }

    [TestMethod]
    public void ClearFormulasTest()
    {
        using ExcelPackage package = new ExcelPackage();
        ExcelWorksheet? worksheet = package.Workbook.Worksheets.Add("Sheet1");
        worksheet.Cells["A1"].Value = 1;
        worksheet.Cells["A2"].Value = 2;
        worksheet.Cells["A3"].Formula = "SUM(A1:A2)";
        worksheet.Cells["A4"].Formula = "SUM(A1:A2)";
        worksheet.Calculate();
        Assert.AreEqual(3d, worksheet.Cells["A3"].Value);
        Assert.AreEqual(3d, worksheet.Cells["A4"].Value);
        Assert.AreEqual("SUM(A1:A2)", worksheet.Cells["A3"].Formula);
        Assert.AreEqual("SUM(A1:A2)", worksheet.Cells["A4"].Formula);
        worksheet.Cells["A3"].ClearFormulas();
        Assert.AreEqual(3d, worksheet.Cells["A3"].Value);
        Assert.AreEqual(string.Empty, worksheet.Cells["A3"].Formula);
        Assert.AreEqual("SUM(A1:A2)", worksheet.Cells["A4"].Formula);
    }

    [TestMethod]
    public void ClearFormulaValuesTest()
    {
        using ExcelPackage package = new ExcelPackage();
        ExcelWorksheet? worksheet = package.Workbook.Worksheets.Add("Sheet1");
        worksheet.Cells["A1"].Value = 1;
        worksheet.Cells["A2"].Value = 2;
        worksheet.Cells["A3"].Formula = "SUM(A1:A2)";
        worksheet.Cells["A4"].Formula = "SUM(A1:A2)";
        worksheet.Calculate();
        Assert.AreEqual(3d, worksheet.Cells["A3"].Value);
        Assert.AreEqual(3d, worksheet.Cells["A4"].Value);
        worksheet.Cells["A3"].ClearFormulaValues();
        Assert.IsNull(worksheet.Cells["A3"].Value);
        Assert.AreEqual(3d, worksheet.Cells["A4"].Value);
    }
}