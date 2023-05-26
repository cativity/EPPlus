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
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup;

[TestClass]
public class IndexTests
{
    private ParsingContext _parsingContext;
    private ExcelPackage _package;
    private ExcelWorksheet _worksheet;

    [TestInitialize]
    public void Initialize()
    {
        this._parsingContext = ParsingContext.Create();
        this._package = new ExcelPackage(new MemoryStream());
        this._worksheet = this._package.Workbook.Worksheets.Add("test");
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._package.Dispose();
    }
        
    [TestMethod]
    public void Index_Should_Return_Value_By_Index()
    {
        Index? func = new Index();
        CompileResult? result = func.Execute(
                                             FunctionsHelper.CreateArgs(
                                                                        FunctionsHelper.CreateArgs(1, 2, 5),
                                                                        3
                                                                       ),
                                             this._parsingContext);
        Assert.AreEqual(5, result.Result);
    }

    [TestMethod]
    public void Index_Should_Handle_SingleRange()
    {
        this._worksheet.Cells["A1"].Value = 1d;
        this._worksheet.Cells["A2"].Value = 3d;
        this._worksheet.Cells["A3"].Value = 5d;

        this._worksheet.Cells["A4"].Formula = "INDEX(A1:A3;3)";

        this._worksheet.Calculate();

        Assert.AreEqual(5d, this._worksheet.Cells["A4"].Value);
    }

    [TestMethod]
    public void Index_SameColumn()
    {
        this._worksheet.Cells["A1"].Value = "value";
        this._worksheet.Cells["A2"].Formula = "INDEX(A:A,1,1)";

        this._worksheet.Calculate();

        Assert.AreEqual("value", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void Index_With_Match()
    {
        this._worksheet.Cells["A1"].Value = "key1";
        this._worksheet.Cells["A2"].Value = "key2";
        this._worksheet.Cells["B1"].Value = "value_to_match";
        this._worksheet.Cells["B2"].Formula = "INDEX($B:$B, MATCH(\"key1\", $A:$A, FALSE), 1)";

        this._worksheet.Calculate();

        Assert.AreEqual("value_to_match", this._worksheet.Cells["B2"].Value);
    }
}