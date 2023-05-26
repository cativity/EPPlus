using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions;
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions;

[TestClass]
public class LoadFromCollectionAttributesComplexTypeTests
{
    private List<Outer> _collection = new List<Outer>();
    private List<OuterWithHeaders> _collectionHeaders = new List<OuterWithHeaders>();
    private List<OuterReversedSortOrder> _collectionReversed = new List<OuterReversedSortOrder>();
    private List<OuterSubclass> _collectionInheritence = new List<OuterSubclass>();

    [TestInitialize]
    public void Initialize()
    {
        this._collection.Add(new Outer
        {
            ApprovedUtc = new DateTime(2021, 7, 1),
            Organization = new Organization 
            { 
                OrgLevel3 = "ABC", 
                OrgLevel4 = "DEF", 
                OrgLevel5 = "GHI"
            },
            Acknowledged = true
        });

        this._collectionHeaders.Add(new OuterWithHeaders
        {
            ApprovedUtc = new DateTime(2021, 7, 1),
            Organization = new Organization
            {
                OrgLevel3 = "ABC",
                OrgLevel4 = "DEF",
                OrgLevel5 = "GHI"
            },
            Acknowledged = true
        });

        this._collectionReversed.Add(new OuterReversedSortOrder
        {
            ApprovedUtc = new DateTime(2021, 7, 1),
            Organization = new OrganizationReversedSortOrder
            {
                OrgLevel3 = "ABC",
                OrgLevel4 = "DEF",
                OrgLevel5 = "GHI"
            },
            Acknowledged = true
        });

        this._collectionInheritence.Add(new OuterSubclass
        {
            ApprovedUtc = new DateTime(2021, 7, 1),
            Organization = new OrganizationSubclass
            {
                OrgLevel3 = "ABC",
                OrgLevel4 = "DEF",
                OrgLevel5 = "GHI"
            },
            Acknowledged = true
        });
    }

    [TestCleanup]
    public void Cleanup()
    {
        this._collection.Clear();
    }

    [TestMethod]
    public void ShouldSetupColumnsWithPath()
    {
        LoadFromCollectionColumns<Outer>? cols = new LoadFromCollectionColumns<Outer>(LoadFromCollectionParams.DefaultBindingFlags, Enumerable.Empty<string>().ToList());
        List<ColumnInfo>? result = cols.Setup();
        Assert.AreEqual(5, result.Count, "List did not contain 5 elements as expected");
        Assert.AreEqual("ApprovedUtc", result[0].Path);
        Assert.AreEqual("Organization.OrgLevel3", result[1].Path);
    }

    [TestMethod]
    public void ShouldSetupColumnsWithPathSorted()
    {
        LoadFromCollectionColumns<OuterReversedSortOrder>? cols = new LoadFromCollectionColumns<OuterReversedSortOrder>(LoadFromCollectionParams.DefaultBindingFlags);
        List<ColumnInfo>? result = cols.Setup();
        Assert.AreEqual(5, result.Count, "List did not contain 5 elements as expected");
        Assert.AreEqual("Acknowledged", result[0].Path);
        Assert.AreEqual("Organization.OrgLevel5", result[1].Path);
        Assert.AreEqual("ApprovedUtc", result.Last().Path);
    }

    [TestMethod]
    public void ShouldSetupColumnsWithPathSortedByClassAttribute()
    {
        List<string>? order = new List<string>
        {
            "ApprovedUtc",
            "Acknowledged",
            "Organization.OrgLevel5"
        };
        LoadFromCollectionColumns<OuterReversedSortOrder>? cols = new LoadFromCollectionColumns<OuterReversedSortOrder>(LoadFromCollectionParams.DefaultBindingFlags, order);
        List<ColumnInfo>? result = cols.Setup();
        Assert.AreEqual(5, result.Count, "List did not contain 5 elements as expected");
        Assert.AreEqual("ApprovedUtc", result[0].Path);
        Assert.AreEqual("Acknowledged", result[1].Path);
        Assert.AreEqual("Organization.OrgLevel5", result[2].Path);
        Assert.AreEqual("Organization.OrgLevel4", result[3].Path);
        Assert.AreEqual("Organization.OrgLevel3", result[4].Path);

    }

    [TestMethod]
    public void ShouldLoadFromComplexTypeMember()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
        ws.Cells["A1"].LoadFromCollection(this._collection);
        Assert.AreEqual("ABC", ws.Cells["B1"].Value);
    }

    [TestMethod]
    public void ShouldLoadFromComplexTypeMemberSorted()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
        ws.Cells["A1"].LoadFromCollection(this._collectionReversed);
        Assert.IsTrue((bool)ws.Cells["A1"].Value);
        Assert.AreEqual("GHI", ws.Cells["B1"].Value);
        Assert.AreEqual(new DateTime(2021, 7, 1), ws.Cells["E1"].Value);
    }

    [TestMethod]
    public void ShouldLoadFromComplexTypeMemberWhenComplexMemberIsNull()
    {
        Outer? obj = this._collection.First();
        obj.Organization = null;
        this._collection[0] = obj;
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
        ws.Cells["A1"].LoadFromCollection(this._collection);
        Assert.IsNull(ws.Cells["B1"].Value);
    }

    [TestMethod]
    public void ShouldLoadFromComplexTypeMemberWhenComplexMemberIsNull_WithHeaders()
    {
        OuterWithHeaders? obj = this._collectionHeaders.First();
        obj.Organization = null;
        this._collectionHeaders[0] = obj;
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
        ws.Cells["A1"].LoadFromCollection(this._collectionHeaders);
        Assert.AreEqual("Org Level 3", ws.Cells["B1"].Value);
        Assert.IsNull(ws.Cells["B2"].Value);
    }

    [TestMethod]
    public void ShouldSetHeaderPrefixOnComplexClassProperty_WithTableColumnAttributeOnChildProperty()
    {
        IEnumerable<IntegratedPlatformExcelRow>? items = ExcelItems.GetItems1();
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
        ws.Cells["A1"].LoadFromCollection(items);
        object? cv = ws.Cells["F1"].Value;
        Assert.AreEqual("Collateral Owner Email", cv);
    }

    [TestMethod]
    public void ShouldSetHeaderPrefixOnComplexClassProperty_WithoutTableColumnAttributeOnChildProperty()
    {
        IEnumerable<IntegratedPlatformExcelRow>? items = ExcelItems.GetItems1();
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
        ws.Cells["A1"].LoadFromCollection(items);
        object? cv = ws.Cells["G1"].Value;
        Assert.AreEqual("Collateral Owner Name", cv);
    }

    [TestMethod]
    public void ShouldLoadFromComplexInheritence()
    {
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
        ws.Cells["A1"].LoadFromCollection(this._collectionInheritence);
        Assert.AreEqual("ABC", ws.Cells["B1"].Value);
    }

    [TestMethod]
    public void LoadComplexTest2()
    {
        using ExcelPackage? package = new ExcelPackage();
        IEnumerable<IntegratedPlatformExcelRow>? items = ExcelItems.GetItems1();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        sheet.Cells["A1"].LoadFromCollection(items);
        Assert.AreEqual("Product Family", sheet.Cells["A1"].Value);
        Assert.AreEqual("PCH Die Name", sheet.Cells["B1"].Value);
        Assert.AreEqual("Collateral Owner Email", sheet.Cells["F1"].Value);
        Assert.AreEqual("Mission Control Lead Email", sheet.Cells["I1"].Value);
        Assert.AreEqual("Created (GMT)", sheet.Cells["L1"].Value);
    }
}