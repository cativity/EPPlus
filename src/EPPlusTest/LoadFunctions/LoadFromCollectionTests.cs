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
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;

namespace EPPlusTest.LoadFunctions;

[TestClass]
public class LoadFromCollectionTests : TestBase
{
    internal abstract class BaseClass
    {
        public string Id { get; set; }

        public string Name { get; set; }
    }

    internal class Implementation : BaseClass
    {
        public int Number { get; set; }
    }

    [System.ComponentModel.Description("The color Red")]
    internal enum Aenum
    {
        [System.ComponentModel.Description("The color Red")]
        Red,

        [System.ComponentModel.Description("The color Blue")]
        Blue,
        Green
    }

    internal class EnumClass
    {
        public int Id { get; set; }

        public Aenum Enum { get; set; }

        [System.ComponentModel.Description("Nullable Enum")]
        public Aenum? NullableEnum { get; set; }
    }

    internal class Aclass
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public int Number { get; set; }
    }

    internal class BClass
    {
        [DisplayName("MyId")]
        public string Id { get; set; }

        [System.ComponentModel.Description("MyName")]
        public string Name { get; set; }

        [EpplusTableColumn(Order = 3)]
        public int Number { get; set; }
    }

    internal class CamelCasedClass
    {
        public string IdOfThisInstance { get; set; }

        public string CamelCased_And_Underscored { get; set; }
    }

    internal class UrlClass : BClass
    {
        [EpplusIgnore]
        public string EMailAddress { get; set; }

        [EpplusTableColumn(Order = 5, Header = "My Mail To")]
        public ExcelHyperLink MailTo
        {
            get
            {
                ExcelHyperLink? url = new ExcelHyperLink("mailto:" + this.EMailAddress);
                url.Display = this.Name;

                return url;
            }
        }

        [EpplusTableColumn(Order = 4)]
        public Uri Url { get; set; }
    }

    [TestMethod]
    public void ShouldNotIncludeHeadersWhenPrintHeadersIsOmitted()
    {
        List<Aclass>? items = new List<Aclass>()
        {
            new Aclass() { Id = "123", Name = "Item 1", Number = 3 }, new Aclass() { Id = "456", Name = "Item 2", Number = 6 }
        };

        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items);

        Assert.AreEqual("123", sheet.Cells["C1"].Value);
        Assert.AreEqual(6, sheet.Cells["E2"].Value);
        Assert.AreEqual(3, sheet.Dimension._fromCol);
        Assert.AreEqual(5, sheet.Dimension._toCol);
        Assert.AreEqual(1, sheet.Dimension._fromRow);
        Assert.AreEqual(2, sheet.Dimension._toRow);
    }

    [TestMethod]
    public void ShouldIncludeHeaders()
    {
        List<Aclass>? items = new List<Aclass>()
        {
            new Aclass() { Id = "123", Name = "Item 1", Number = 3 }, new Aclass() { Id = "456", Name = "Item 2", Number = 6 }
        };

        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items, true);
        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldIncludeHeadersAndTableStyle()
    {
        List<Aclass>? items = new List<Aclass>()
        {
            new Aclass() { Id = "123", Name = "Item 1", Number = 3 }, new Aclass() { Id = "456", Name = "Item 2", Number = 6 }
        };

        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);
        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldUseAclassProperties()
    {
        List<Aclass>? items = new List<Aclass>() { new Aclass() { Id = "123", Name = "Item 1", Number = 3 } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
        Assert.AreEqual("123", sheet.Cells["C2"].Value);
    }

    [TestMethod]
    public void ShouldUseDisplayNameAttribute()
    {
        List<BClass>? items = new List<BClass>() { new BClass() { Id = "123", Name = "Item 1", Number = 3 } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

        Assert.AreEqual("MyId", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldFilterMembers()
    {
        List<BaseClass>? items = new List<BaseClass>() { new Implementation() { Id = "123", Name = "Item 1", Number = 3 } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        Type? t = typeof(Implementation);

        _ = sheet.Cells["C1"]
             .LoadFromCollection(items,
                                 true,
                                 TableStyles.Dark1,
                                 LoadFromCollectionParams.DefaultBindingFlags,
                                 new MemberInfo[] { t.GetProperty("Id"), t.GetProperty("Name") });

        Assert.AreEqual(1, sheet.Dimension._toCol - sheet.Dimension._fromCol);
        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
        Assert.AreEqual("Name", sheet.Cells["D1"].Value);
        Assert.IsNull(sheet.Cells["E1"].Value);
        Assert.AreEqual("123", sheet.Cells["C2"].Value);
        Assert.AreEqual("Item 1", sheet.Cells["D2"].Value);
        Assert.IsNull(sheet.Cells["E2"].Value);
    }

    [TestMethod]
    public void ShouldFilterOneMember()
    {
        List<BaseClass>? items = new List<BaseClass>() { new Implementation() { Id = "123", Name = "Item 1", Number = 3 } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        Type? t = typeof(Implementation);

        _ = sheet.Cells["C1"]
             .LoadFromCollection(items, true, TableStyles.Dark1, LoadFromCollectionParams.DefaultBindingFlags, new MemberInfo[] { t.GetProperty("Id"), });

        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
        Assert.AreEqual("123", sheet.Cells["C2"].Value);
    }

    [TestMethod]
    public void ShouldUseDescriptionAttribute()
    {
        List<BClass>? items = new List<BClass>() { new BClass() { Id = "123", Name = "Item 1", Number = 3 } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

        Assert.AreEqual("MyName", sheet.Cells["D1"].Value);
    }

    [TestMethod]
    public void ShouldUseBaseClassProperties()
    {
        List<BaseClass>? items = new List<BaseClass>() { new Implementation() { Id = "123", Name = "Item 1", Number = 3 } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldUseAnonymousProperties()
    {
        List<BaseClass>? objs = new List<BaseClass>() { new Implementation() { Id = "123", Name = "Item 1", Number = 3 } };
        var items = objs.Select(x => new { Id = x.Id, Name = x.Name }).ToList();
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidCastException))]
    public void ShouldThrowInvalidCastExceptionIf()
    {
        List<BaseClass>? objs = new List<BaseClass>() { new Implementation() { Id = "123", Name = "Item 1", Number = 3 } };
        var items = objs.Select(x => new { Id = x.Id, Name = x.Name }).ToList();
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");
        _ = sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1, BindingFlags.Public | BindingFlags.Instance, typeof(string).GetMembers());

        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldUseLambdaConfig()
    {
        List<Aclass>? items = new List<Aclass>() { new Aclass() { Id = "123", Name = "Item 1", Number = 3 } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");

        _ = sheet.Cells["C1"]
             .LoadFromCollection(items,
                                 c =>
                                 {
                                     c.PrintHeaders = true;
                                     c.TableStyle = TableStyles.Dark1;
                                 });

        Assert.AreEqual("Id", sheet.Cells["C1"].Value);
        Assert.AreEqual("123", sheet.Cells["C2"].Value);
        Assert.AreEqual(3, sheet.Cells["E2"].Value);
        Assert.AreEqual(1, sheet.Tables.Count());
    }

    [TestMethod]
    public void ShouldParseCamelCasedHeaders()
    {
        List<CamelCasedClass>? items = new List<CamelCasedClass>() { new CamelCasedClass() { IdOfThisInstance = "123" } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");

        _ = sheet.Cells["C1"]
             .LoadFromCollection(items,
                                 c =>
                                 {
                                     c.PrintHeaders = true;
                                     c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace;
                                 });

        Assert.AreEqual("Id Of This Instance", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldParseCamelCasedAndUnderscoredHeaders()
    {
        List<CamelCasedClass>? items = new List<CamelCasedClass>() { new CamelCasedClass() { CamelCased_And_Underscored = "123" } };
        using ExcelPackage? pck = new ExcelPackage(new MemoryStream());
        ExcelWorksheet? sheet = pck.Workbook.Worksheets.Add("sheet");

        _ = sheet.Cells["C1"]
             .LoadFromCollection(items,
                                 c =>
                                 {
                                     c.PrintHeaders = true;
                                     c.HeaderParsingType = HeaderParsingTypes.UnderscoreAndCamelCaseToSpace;
                                 });

        Assert.AreEqual("Camel Cased And Underscored", sheet.Cells["D1"].Value);
    }

    [TestMethod]
    public void ShouldLoadExpandoObjects()
    {
        dynamic o1 = new ExpandoObject();
        o1.Id = 1;
        o1.Name = "TestName 1";
        dynamic o2 = new ExpandoObject();
        o2.Id = 2;
        o2.Name = "TestName 2";
        List<ExpandoObject>? items = new List<ExpandoObject>() { o1, o2 };
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.None);

        Assert.AreEqual("Id", sheet.Cells["A1"].Value);
        Assert.AreEqual(1, sheet.Cells["A2"].Value);
        Assert.AreEqual("TestName 2", sheet.Cells["B3"].Value);
    }

    [TestMethod]
    public void ShouldSetHyperlinkForURIs()
    {
        List<UrlClass>? items = new List<UrlClass>()
        {
            new UrlClass { Id = "1", Name = "Person 1", EMailAddress = "person1@somewhe.re" },
            new UrlClass { Id = "2", Name = "Person 2", EMailAddress = "person2@somewhe.re" },
            new UrlClass { Id = "2", Name = "Person with Url", EMailAddress = "person2@somewhe.re", Url = new Uri("https://epplussoftware.com") },
        };

        using ExcelPackage? package = OpenPackage("LoadFromCollectionUrls.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1);

        Assert.AreEqual("MyId", sheet.Cells["A1"].Value);
        Assert.AreEqual("MyName", sheet.Cells["B1"].Value);
        Assert.AreEqual("Number", sheet.Cells["C1"].Value);
        Assert.AreEqual("Url", sheet.Cells["D1"].Value);
        Assert.AreEqual("My Mail To", sheet.Cells["E1"].Value);

        Assert.AreEqual("1", sheet.Cells["A2"].Value);
        Assert.AreEqual("Person 2", sheet.Cells["B3"].Value);
        Assert.IsInstanceOfType(sheet.Cells["E3"].Hyperlink, typeof(ExcelHyperLink));
        Assert.AreEqual("Person 2", sheet.Cells["E3"].Value);

        SaveAndCleanup(package);
    }

    [TestMethod]
    public void LoadListOfEnumWithDescription()
    {
        List<Aenum>? items = new List<Aenum>() { Aenum.Red, Aenum.Green, Aenum.Blue };

        using ExcelPackage? package = OpenPackage("LoadFromCollectionEnumDescrAtt.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("EnumList");
        _ = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1);
        Assert.AreEqual("The color Red", sheet.Cells["A1"].Value);
        Assert.AreEqual("Green", sheet.Cells["A2"].Value);
        Assert.AreEqual("The color Blue", sheet.Cells["A3"].Value);
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void LoadListOfNullableEnumWithDescription()
    {
        List<Aenum?>? items = new List<Aenum?>() { Aenum.Red, Aenum.Green, Aenum.Blue };

        using ExcelPackage? package = OpenPackage("LoadFromCollectionNullableEnumDescrAtt.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("NullableEnumList");
        _ = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1);
        Assert.AreEqual("The color Red", sheet.Cells["A1"].Value);
        Assert.AreEqual("Green", sheet.Cells["A2"].Value);
        Assert.AreEqual("The color Blue", sheet.Cells["A3"].Value);
        SaveAndCleanup(package);
    }

    [TestMethod]
    public void LoadListOfClassWithEnumWithDescription()
    {
        List<EnumClass>? items = new List<EnumClass>()
        {
            new EnumClass() { Id = 1, Enum = Aenum.Red, NullableEnum = Aenum.Blue },
            new EnumClass() { Id = 2, Enum = Aenum.Blue, NullableEnum = null },
            new EnumClass() { Id = 3, Enum = Aenum.Green, NullableEnum = Aenum.Red },
        };

        using ExcelPackage? package = OpenPackage("LoadFromCollectionClassWithEnumDescrAtt.xlsx", true);
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1);
        Assert.AreEqual("Id", sheet.Cells["A1"].Value);
        Assert.AreEqual("Enum", sheet.Cells["B1"].Value);
        Assert.AreEqual("Nullable Enum", sheet.Cells["C1"].Value);
        Assert.AreEqual(1, sheet.Cells["A2"].Value);
        Assert.AreEqual("The color Red", sheet.Cells["B2"].Value);
        Assert.AreEqual("The color Blue", sheet.Cells["C2"].Value);
        Assert.AreEqual(2, sheet.Cells["A3"].Value);
        Assert.AreEqual("The color Blue", sheet.Cells["B3"].Value);
        Assert.IsNull(sheet.Cells["C3"].Value);
        Assert.AreEqual(3, sheet.Cells["A4"].Value);
        Assert.AreEqual("Green", sheet.Cells["B4"].Value);
        Assert.AreEqual("The color Red", sheet.Cells["C4"].Value);

        SaveAndCleanup(package);
    }
}