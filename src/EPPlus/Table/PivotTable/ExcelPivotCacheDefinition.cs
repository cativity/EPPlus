/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using System;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;
using System.Security;
using System.Linq;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// Cache definition. This class defines the source data. Note that one cache definition can be shared between many pivot tables.
/// </summary>
public class ExcelPivotCacheDefinition
{
    ExcelWorkbook _wb;
    internal PivotTableCacheInternal _cacheReference;
    XmlNamespaceManager _nsm;

    internal ExcelPivotCacheDefinition(XmlNamespaceManager nsm, ExcelPivotTable pivotTable)
    {
        this.Relationship = pivotTable.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheDefinition").FirstOrDefault();
        Uri? cacheDefinitionUri = UriHelper.ResolvePartUri(this.Relationship.SourceUri, this.Relationship.TargetUri);
        this.PivotTable = pivotTable;
        this._wb = pivotTable.WorkSheet.Workbook;
        this._nsm = nsm;

        ExcelWorkbook.PivotTableCacheRangeInfo? c =
            this._wb._pivotTableCaches.Values.FirstOrDefault(x => x.PivotCaches.Exists(y => y.CacheDefinitionUri.OriginalString
                                                                                            == cacheDefinitionUri.OriginalString));

        if (c == null)
        {
            if (this._wb._pivotTableIds.ContainsKey(cacheDefinitionUri))
            {
                int cid = this._wb._pivotTableIds[cacheDefinitionUri];
                this._cacheReference = new PivotTableCacheInternal(this._wb, cacheDefinitionUri, cid);
                this._wb.AddPivotTableCache(this._cacheReference, false);
            }
            else
            {
                throw new Exception("Internal error: Pivot table uri does not exist in package.");
            }
        }
        else
        {
            this._cacheReference = c.PivotCaches.FirstOrDefault(x => x.CacheDefinitionUri.OriginalString == cacheDefinitionUri.OriginalString);
        }

        this._cacheReference._pivotTables.Add(pivotTable);
    }

    internal ExcelPivotCacheDefinition(XmlNamespaceManager nsm, ExcelPivotTable pivotTable, ExcelRangeBase sourceRange)
    {
        this.PivotTable = pivotTable;
        this._wb = this.PivotTable.WorkSheet.Workbook;
        this._nsm = nsm;
        this._cacheReference = new PivotTableCacheInternal(nsm, this._wb);
        this._cacheReference.InitNew(pivotTable, sourceRange, null);
        this._wb.AddPivotTableCache(this._cacheReference);

        this.Relationship = pivotTable.Part.CreateRelationship(UriHelper.ResolvePartUri(pivotTable.PivotTableUri, this._cacheReference.CacheDefinitionUri),
                                                               TargetMode.Internal,
                                                               ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
    }

    internal ExcelPivotCacheDefinition(XmlNamespaceManager nsm, ExcelPivotTable pivotTable, PivotTableCacheInternal cache)
    {
        this.PivotTable = pivotTable;
        this._wb = this.PivotTable.WorkSheet.Workbook;
        this._nsm = nsm;

        if (cache._wb != this._wb)
        {
            throw new InvalidOperationException("The pivot table and the cache must be in the same workbook.");
        }

        this._cacheReference = cache;
        this._cacheReference._pivotTables.Add(pivotTable);

        _ = pivotTable.Part.CreateRelationship(UriHelper.ResolvePartUri(pivotTable.PivotTableUri, this._cacheReference.CacheDefinitionUri),
                                           TargetMode.Internal,
                                           ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
    }

    internal void Refresh()
    {
        this._cacheReference.RefreshFields();
    }

    internal ZipPackagePart Part { get; set; }

    /// <summary>
    /// Provides access to the XML data representing the cache definition in the package.
    /// </summary>
    public XmlDocument CacheDefinitionXml
    {
        get { return this._cacheReference.CacheDefinitionXml; }
    }

    /// <summary>
    /// The package internal URI to the pivottable cache definition Xml Document.
    /// </summary>
    public Uri CacheDefinitionUri
    {
        get { return this._cacheReference.CacheDefinitionUri; }
    }

    internal ZipPackageRelationship Relationship { get; set; }

    /// <summary>
    /// Referece to the PivotTable object
    /// </summary>
    public ExcelPivotTable PivotTable { get; private set; }

    const string _sourceWorksheetPath = "d:cacheSource/d:worksheetSource/@sheet";
    internal const string _sourceNamePath = "d:cacheSource/d:worksheetSource/@name";
    internal const string _sourceAddressPath = "d:cacheSource/d:worksheetSource/@ref";
    internal ExcelRangeBase _sourceRange = null;

    /// <summary>
    /// The source data range when the pivottable has a worksheet datasource. 
    /// The number of columns in the range must be intact if this property is changed.
    /// The range must be in the same workbook as the pivottable.
    /// </summary>
    public ExcelRangeBase SourceRange
    {
        get { return this._cacheReference.SourceRange; }
        set
        {
            if (this.PivotTable.WorkSheet.Workbook != value.Worksheet.Workbook)
            {
                throw new ArgumentException("Range must be in the same package as the pivottable");
            }

            ExcelRangeBase? sr = this.SourceRange;

            if (value.End.Column - value.Start.Column != sr.End.Column - sr.Start.Column)
            {
                throw new ArgumentException("Cannot change the number of columns(fields) in the SourceRange");
            }

            if (value.FullAddress == this.SourceRange.FullAddress)
            {
                return; //Same
            }

            if (this._wb.GetPivotCacheFromAddress(value.FullAddress, out PivotTableCacheInternal cache))
            {
                _ = this._cacheReference._pivotTables.Remove(this.PivotTable);
                cache._pivotTables.Add(this.PivotTable);
                this._cacheReference = cache;
                this.PivotTable.CacheId = this._cacheReference.CacheId;
                this.Relationship.TargetUri = cache.CacheDefinitionUri;
            }
            else if (this._cacheReference._pivotTables.Count == 1)
            {
                string sourceName = this.SourceRange.GetName();

                if (string.IsNullOrEmpty(sourceName))
                {
                    this._cacheReference.SetXmlNodeString(_sourceWorksheetPath, value.Worksheet.Name);
                    this._cacheReference.SetXmlNodeString(_sourceAddressPath, value.FirstAddress);
                }
                else
                {
                    this._cacheReference.SetXmlNodeString(_sourceNamePath, sourceName);
                }

                this._sourceRange = value;
            }
            else
            {
                _ = this._cacheReference._pivotTables.Remove(this.PivotTable);
                XmlDocument? xml = this._cacheReference.CacheDefinitionXml;
                this._cacheReference = new PivotTableCacheInternal(this._nsm, this._wb);
                this._cacheReference.InitNew(this.PivotTable, value, xml.InnerXml);
                this.PivotTable.CacheId = this._cacheReference.CacheId;
                this._wb.AddPivotTableCache(this._cacheReference);
                this.Relationship.TargetUri = this._cacheReference.CacheDefinitionUri;
            }
        }
    }

    /// <summary>
    /// If Excel will save the source data with the pivot table.
    /// </summary>
    public bool SaveData
    {
        get { return this._cacheReference.SaveData; }
        set { this._cacheReference.SaveData = value; }
    }

    /// <summary>
    /// Type of source data
    /// </summary>
    public eSourceType CacheSource
    {
        get { return this._cacheReference.CacheSource; }
    }
}