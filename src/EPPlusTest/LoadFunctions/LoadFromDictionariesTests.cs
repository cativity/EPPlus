﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions;

[TestClass]
public class LoadFromDictionariesTests
{
    [TestInitialize]
    public void Initialize() =>
        this._items = new List<IDictionary<string, object>>()
        {
            new Dictionary<string, object>() { { "Id", 1 }, { "Name", "TestName 1" } },
            new Dictionary<string, object>() { { "Id", 2 }, { "Name", "TestName 2" } }
        };

    private IEnumerable<IDictionary<string, object>> _items;

    [TestMethod]
    public void ShouldLoadDictionaryWithoutHeaders()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromDictionaries(this._items);

        Assert.AreEqual(1, sheet.Cells["A1"].Value);
        Assert.AreEqual("TestName 2", sheet.Cells["B2"].Value);
    }

    [TestMethod]
    public void ShouldLoadDictionaryWithHeaders()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromDictionaries(this._items, true, TableStyles.None, null);

        Assert.AreEqual("Id", sheet.Cells["A1"].Value);
        Assert.AreEqual(1, sheet.Cells["A2"].Value);
        Assert.AreEqual("TestName 2", sheet.Cells["B3"].Value);
    }

    [TestMethod]
    public void ShouldLoadDictionaryWithParsedHeaders_Default()
    {
        foreach (IDictionary<string, object>? item in this._items)
        {
            item["First_name"] = "test";
        }

        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromDictionaries(this._items, true, TableStyles.None, null);

        Assert.AreEqual("First name", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldLoadDictionaryWithParsedHeaders_CamelCase()
    {
        foreach (IDictionary<string, object>? item in this._items)
        {
            item["FirstName"] = "test";
        }

        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

        _ = sheet.Cells["A1"]
             .LoadFromDictionaries(this._items,
                                   c =>
                                   {
                                       c.PrintHeaders = true;
                                       c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace;
                                   });

        Assert.AreEqual("First Name", sheet.Cells["C1"].Value);
    }

    [TestMethod]
    public void ShouldLoadDictionaryWithHeadersAndTable()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromDictionaries(this._items, true, TableStyles.Dark1, null);

        Assert.AreEqual(1, sheet.Tables.Count);
        Assert.AreEqual(TableStyles.Dark1, sheet.Tables.First().TableStyle);
    }

    [TestMethod]
    public void ShouldLoadDictionaryWithKeysFilter()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromDictionaries(this._items, false, TableStyles.None, new string[] { "Name" });

        Assert.AreEqual("TestName 1", sheet.Cells["A1"].Value);
    }

    [TestMethod]
    public void ShouldLoadDictionaryWithKeysFilterLambdaVersion()
    {
        foreach (IDictionary<string, object>? item in this._items)
        {
            item["Number"] = 1;
        }

        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");

        _ = sheet.Cells["A1"]
             .LoadFromDictionaries(this._items,
                                   c =>
                                   {
                                       c.PrintHeaders = false;
                                       c.TableStyle = TableStyles.None;
                                       c.SetKeys("Name", "Number");
                                   });

        Assert.AreEqual("TestName 1", sheet.Cells["A1"].Value);
        Assert.AreEqual(1, sheet.Cells["B1"].Value);
        Assert.IsNull(sheet.Cells["C1"].Value);
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
        _ = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None, null);

        Assert.AreEqual("Id", sheet.Cells["A1"].Value);
        Assert.AreEqual(1, sheet.Cells["A2"].Value);
        Assert.AreEqual("TestName 2", sheet.Cells["B3"].Value);
    }

    [TestMethod]
    public void ShouldLoadDynamicObjects()
    {
        dynamic o1 = new { Id = 1, Name = "TestName 1" };
        dynamic o2 = new { Id = 2, Name = "TestName 2" };
        List<dynamic>? items = new List<dynamic>() { o1, o2 };
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        _ = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None, null);

        Assert.AreEqual("Id", sheet.Cells["A1"].Value);
        Assert.AreEqual(1, sheet.Cells["A2"].Value);
        Assert.AreEqual("TestName 2", sheet.Cells["B3"].Value);
    }
}