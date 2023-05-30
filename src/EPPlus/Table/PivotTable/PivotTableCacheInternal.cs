using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// Handles the pivot table cache.
/// </summary>
internal class PivotTableCacheInternal : XmlHelper
{
    internal List<ExcelPivotTable> _pivotTables = new List<ExcelPivotTable>();
    internal readonly ExcelWorkbook _wb;

    public PivotTableCacheInternal(XmlNamespaceManager nsm, ExcelWorkbook wb)
        : base(nsm)
    {
        this._wb = wb;
    }

    public PivotTableCacheInternal(ExcelWorkbook wb, Uri uri, int cacheId)
        : base(wb.NameSpaceManager)
    {
        this._wb = wb;
        this.CacheDefinitionUri = uri;
        this.Part = wb._package.ZipPackage.GetPart(uri);

        this.CacheDefinitionXml = new XmlDocument();
        LoadXmlSafe(this.CacheDefinitionXml, this.Part.GetStream());
        this.TopNode = this.CacheDefinitionXml.DocumentElement;

        if (this.CacheId <= 0) //Check if the is set via exLst (used by for example slicers), otherwise set it to the cacheId
        {
            this.CacheId = cacheId;
        }

        ZipPackageRelationship rel = this.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheRecords").FirstOrDefault();

        if (rel != null)
        {
            this.CacheRecordUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
        }

        this._wb.SetNewPivotCacheId(cacheId);
    }

    internal const string _sourceWorksheetPath = "d:cacheSource/d:worksheetSource/@sheet";
    internal const string _sourceNamePath = "d:cacheSource/d:worksheetSource/@name";
    internal const string _sourceAddressPath = "d:cacheSource/d:worksheetSource/@ref";

    internal string Ref
    {
        get { return this.GetXmlNodeString(_sourceAddressPath); }
    }

    internal string SourceName
    {
        get { return this.GetXmlNodeString(_sourceNamePath); }
    }

    internal ExcelRangeBase SourceRange
    {
        get
        {
            ExcelRangeBase sourceRange = null;

            if (this.CacheSource == eSourceType.Worksheet)
            {
                ExcelWorksheet? ws = this._wb.Worksheets[this.GetXmlNodeString(_sourceWorksheetPath)];

                if (ws == null) //Not worksheet, check name or table name
                {
                    string? name = this.GetXmlNodeString(_sourceNamePath);

                    foreach (ExcelNamedRange? n in this._wb.Names)
                    {
                        if (name.Equals(n.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            sourceRange = n;

                            return sourceRange;
                        }
                    }

                    foreach (ExcelWorksheet? w in this._wb.Worksheets)
                    {
                        sourceRange = GetRangeByName(w, name);

                        if (sourceRange != null)
                        {
                            break;
                        }
                    }
                }
                else
                {
                    string? address = this.Ref;

                    if (string.IsNullOrEmpty(address))
                    {
                        string? name = this.SourceName;
                        sourceRange = GetRangeByName(ws, name);
                    }
                    else
                    {
                        sourceRange = ws.Cells[address];
                    }
                }
            }
            else
            {
                throw new ArgumentException("The cache source is not a worksheet");
            }

            return sourceRange;
        }
    }

    private static ExcelRangeBase GetRangeByName(ExcelWorksheet w, string name)
    {
        if (w is ExcelChartsheet)
        {
            return null;
        }

        if (w.Tables._tableNames.ContainsKey(name))
        {
            ExcelTable? t = w.Tables[name];
            int toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;

            return w.Cells[t.Address._fromRow, t.Address._fromCol, toRow, t.Address._toCol];
        }

        foreach (ExcelNamedRange? n in w.Names)
        {
            if (name.Equals(n.Name, StringComparison.OrdinalIgnoreCase))
            {
                return n;
            }
        }

        return null;
    }

    /// <summary>
    /// Reference to the internal package part
    /// </summary>
    internal ZipPackagePart Part { get; set; }

    /// <summary>
    /// Provides access to the XML data representing the cache definition in the package.
    /// </summary>
    internal XmlDocument CacheDefinitionXml { get; set; }

    /// <summary>
    /// The package internal URI to the pivottable cache definition Xml Document.
    /// </summary>
    internal Uri CacheDefinitionUri { get; set; }

    internal Uri CacheRecordUri { get; set; }

    internal ZipPackageRelationship RecordRelationship { get; set; }

    internal string RecordRelationshipId
    {
        get { return this.GetXmlNodeString("@r:id"); }
        set { this.SetXmlNodeString("@r:id", value, true); }
    }

    List<ExcelPivotTableCacheField> _fields;

    internal List<ExcelPivotTableCacheField> Fields
    {
        get
        {
            if (this._fields == null)
            {
                this.LoadFields();

                //RefreshFields();
            }

            return this._fields;
        }
    }

    private void LoadFields()
    {
        //Add fields.
        int index = 0;
        this._fields = new List<ExcelPivotTableCacheField>();

        foreach (XmlNode node in this.CacheDefinitionXml.DocumentElement.SelectNodes("d:cacheFields/d:cacheField", this.NameSpaceManager))
        {
            this._fields.Add(new ExcelPivotTableCacheField(this.NameSpaceManager, node, this, index++));
        }
    }

    internal void RefreshFields()
    {
        List<List<string>>? tableFields = this.GetTableFields();
        List<ExcelPivotTableCacheField>? fields = new List<ExcelPivotTableCacheField>();
        ExcelRangeBase? r = this.SourceRange;
        bool cacheUpdated = false;

        for (int col = r._fromCol; col <= r._toCol; col++)
        {
            int ix = col - r._fromCol;

            if (this._fields != null && ix < this._fields.Count && this._fields[ix].Grouping != null)
            {
                fields.Add(this._fields[ix]);
            }
            else
            {
                ExcelWorksheet? ws = r.Worksheet;
                string? name = ws.GetValue(r._fromRow, col)?.ToString().Trim();
                ExcelPivotTableCacheField field;

                if (this._fields == null || ix >= this._fields?.Count)
                {
                    if (string.IsNullOrEmpty(name))
                    {
                        throw new
                            InvalidOperationException($"Pivot Cache with id {this.CacheId} is invalid . Contains reference to a column with an empty header");
                    }

                    field = this.CreateField(name, ix);
                    field.TopNode.InnerXml = "<sharedItems/>";

                    foreach (ExcelPivotTable? pt in this._pivotTables)
                    {
                        _ = pt.Fields.AddField(ix);
                    }

                    cacheUpdated = true;
                }
                else
                {
                    field = this._fields[ix];
                    field.SharedItems.Clear();

                    if (cacheUpdated == false && string.IsNullOrEmpty(name) == false && !field.Name.StartsWith(name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        cacheUpdated = true;
                    }
                }

                if (!string.IsNullOrEmpty(name) && !field.Name.StartsWith(name))
                {
                    field.Name = name;
                }

                //var hs = new HashSet<object>();
                //var dimensionToRow = ws.Dimension?._toRow ?? r._fromRow + 1;
                //var toRow = r._toRow < dimensionToRow ? r._toRow : dimensionToRow;
                //for (int row = r._fromRow + 1; row <= toRow; row++)
                //{
                //    ExcelPivotTableCacheField.AddSharedItemToHashSet(hs, ws.GetValue(row, col));
                //}
                //field.SharedItems._list = hs.ToList();
                fields.Add(field);
            }
        }

        if (this._fields != null)
        {
            for (int i = fields.Count; i < this._fields.Count; i++)
            {
                fields.Add(this._fields[i]);
            }
        }

        this._fields = fields;

        if (r.Columns != fields.Count)
        {
            this.RemoveDeletedFields(r);
        }

        if (cacheUpdated)
        {
            this.UpdateRowColumnPageFields(tableFields);
        }

        this.RefreshPivotTableItems();
    }

    //private void SyncFields(List<ExcelPivotTableCacheField> fields)
    //{

    //}

    //private void SyncFields()
    //{
    //    var r = SourceRange;
    //    foreach(var pt in _pivotTables)
    //    {       
    //        var newList = new List<ExcelPivotTableField>();
    //        foreach (var f in pt.Fields)
    //        {                    
    //            if (pt.CacheDefinition._cacheReference.Fields.Any(x=>x.Name.Equals(f.Name))
    //            {
    //                f.TopNode.RemoveChild(f.TopNode);                            
    //            }
    //            else
    //            {
    //                newList.Add(f);
    //            }
    //        }
    //        pt.Fields._list = newList;
    //    }
    //}

    private void RemoveDeletedFields(ExcelRangeBase r)
    {
        for (int i = 0; i < this._pivotTables.Count; i++)
        {
            ExcelPivotTable? pt = this._pivotTables[i];

            for (int p = r.Columns; p < pt.Fields.Count; p++)
            {
                if (pt.Fields[p].Cache.DatabaseField)
                {
                    pt.Fields.RemoveAt(pt.Fields.Count - 1);
                    p--;
                }
            }
        }

        int calcFields = 0;

        while (r.Columns + calcFields < this._fields.Count)
        {
            ExcelPivotTableCacheField? f = this._fields[this._fields.Count - 1];

            if (f.DatabaseField)
            {
                _ = f.TopNode.ParentNode.RemoveChild(f.TopNode);
                _ = this._fields.Remove(f);
            }
            else
            {
                calcFields++;
            }
        }
    }

    private void UpdateRowColumnPageFields(List<List<string>> tableFields)
    {
        for (int tblIx = 0; tblIx < this._pivotTables.Count; tblIx++)
        {
            List<string>? l = tableFields[tblIx];
            ExcelPivotTable? tbl = this._pivotTables[tblIx];

            tbl.PageFields._list.ForEach(x =>
            {
                x.IsPageField = false;
                x.Axis = ePivotFieldAxis.None;
            });

            tbl.ColumnFields._list.ForEach(x =>
            {
                x.IsColumnField = false;
                x.Axis = ePivotFieldAxis.None;
            });

            tbl.RowFields._list.ForEach(x =>
            {
                x.IsRowField = false;
                x.Axis = ePivotFieldAxis.None;
            });

            tbl.DataFields._list.ForEach(x =>
            {
                x.Field.IsDataField = false;
                x.Field.Axis = ePivotFieldAxis.None;
            });

            this.ChangeIndex(tbl.PageFields, l);
            this.ChangeIndex(tbl.ColumnFields, l);
            this.ChangeIndex(tbl.RowFields, l);

            for (int i = 0; i < tbl.DataFields.Count; i++)
            {
                ExcelPivotTableDataField? df = tbl.DataFields[i];
                string? prevName = l[df.Index];
                int newIx = this._fields.FindIndex(x => x.Name.Equals(prevName, StringComparison.CurrentCultureIgnoreCase));

                if (newIx >= 0)
                {
                    df.Index = newIx;
                    df.Field = tbl.Fields[newIx];
                    df.Field.IsDataField = true;
                }
                else
                {
                    tbl.DataFields._list.RemoveAt(i--);
                }

                foreach (ExcelPivotTableAreaStyle s in tbl.Styles)
                {
                    if (s.FieldIndex == df.Index)
                    {
                        s.FieldIndex = newIx;
                    }

                    foreach (ExcelPivotAreaReference c in s.Conditions.Fields)
                    {
                        if (c.FieldIndex == df.Index)
                        {
                            c.FieldIndex = newIx;
                        }
                    }

                    if (s.Conditions.DataFields.FieldIndex == df.Index)
                    {
                        s.Conditions.DataFields.FieldIndex = newIx;
                    }
                }
            }
        }
    }

    private void ChangeIndex(ExcelPivotTableRowColumnFieldCollection fields, List<string> prevFields)
    {
        List<ExcelPivotTableField>? newFields = new List<ExcelPivotTableField>();

        for (int i = 0; i < fields.Count; i++)
        {
            ExcelPivotTableField? f = fields[i];
            string? prevName = prevFields[f.Index];
            int ix = this._fields.FindIndex(x => x.Name.Equals(prevName, StringComparison.CurrentCultureIgnoreCase));

            if (ix >= 0)
            {
                ExcelPivotTableField? fld = fields._table.Fields[ix];

                newFields.Add(fld);

                if (fld.PageFieldSettings != null)
                {
                    fld.PageFieldSettings.Index = ix;
                    //fld.PageFieldSettings._field = fld;
                }

                foreach (ExcelPivotTableAreaStyle s in f._pivotTable.Styles)
                {
                    if (s.FieldIndex == f.Index)
                    {
                        s.FieldIndex = ix;
                    }

                    foreach (ExcelPivotAreaReference c in s.Conditions.Fields)
                    {
                        if (c.FieldIndex == f.Index)
                        {
                            c.FieldIndex = ix;
                        }
                    }
                }
            }
        }

        fields.Clear();
        newFields.ForEach(x => fields.Add(x));
    }

    private List<List<string>> GetTableFields()
    {
        List<List<string>>? tableFields = new List<List<string>>();

        foreach (ExcelPivotTable? tbl in this._pivotTables)
        {
            List<string>? l = new List<string>();
            tableFields.Add(l);

            foreach (ExcelPivotTableField? field in tbl.Fields.OrderBy(x => x.Index))
            {
                l.Add(field.Name.ToLower());
            }
        }

        return tableFields;
    }

    private void RefreshPivotTableItems()
    {
        foreach (ExcelPivotTable? pt in this._pivotTables)
        {
            if (pt.CacheDefinition.CacheSource == eSourceType.Worksheet)
            {
                int fieldCount = Math.Min(pt.CacheDefinition.SourceRange.Columns, pt.Fields.Count);

                for (int i = 0; i < fieldCount; i++)
                {
                    pt.Fields[i].Items.Refresh();
                }
            }
        }
    }

    internal eSourceType CacheSource
    {
        get
        {
            string? s = this.GetXmlNodeString("d:cacheSource/@type");

            if (s == "")
            {
                return eSourceType.Worksheet;
            }
            else
            {
                return (eSourceType)Enum.Parse(typeof(eSourceType), s, true);
            }
        }
    }

    internal void InitNew(ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, string xml)
    {
        ZipPackage? pck = pivotTable.WorkSheet._package.ZipPackage;

        this.CacheDefinitionXml = new XmlDocument();
        ExcelWorksheet? sourceWorksheet = pivotTable.WorkSheet.Workbook.Worksheets[sourceAddress.WorkSheetName];

        if (xml == null)
        {
            LoadXmlSafe(this.CacheDefinitionXml, GetStartXml(sourceWorksheet, sourceAddress), Encoding.UTF8);
            this.TopNode = this.CacheDefinitionXml.DocumentElement;
        }
        else
        {
            this.CacheDefinitionXml = new XmlDocument();
            this.CacheDefinitionXml.LoadXml(xml);
            this.TopNode = this.CacheDefinitionXml.DocumentElement;

            string sourceName = this.SourceRange.GetName();

            if (string.IsNullOrEmpty(sourceName))
            {
                this.SetXmlNodeString(_sourceWorksheetPath, sourceAddress.WorkSheetName);
                this.SetXmlNodeString(_sourceAddressPath, sourceAddress.Address);
            }
            else
            {
                this.SetXmlNodeString(_sourceNamePath, sourceName);
            }
        }

        this.CacheId = this._wb.GetNewPivotCacheId();

        int c = this.CacheId;
        this.CacheDefinitionUri = GetNewUri(pck, "/xl/pivotCache/pivotCacheDefinition{0}.xml", ref c);
        this.Part = pck.CreatePart(this.CacheDefinitionUri, ContentTypes.contentTypePivotCacheDefinition);

        this.AddRecordsXml();
        this.LoadFields();
        this.CacheDefinitionXml.Save(this.Part.GetStream());
        this._pivotTables.Add(pivotTable);
    }

    internal void ResetRecordXml(ZipPackage pck)
    {
        if (this.CacheRecordUri == null)
        {
            return;
        }

        XmlDocument? cacheRecord = new XmlDocument();
        cacheRecord.LoadXml("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");
        ZipPackagePart recPart;

        if (pck.PartExists(this.CacheRecordUri))
        {
            recPart = pck.GetPart(this.CacheRecordUri);
        }
        else
        {
            recPart = pck.CreatePart(this.CacheRecordUri, ContentTypes.contentTypePivotCacheRecords);
        }

        cacheRecord.Save(recPart.GetStream(FileMode.Create, FileAccess.Write));
    }

    private static string GetStartXml(ExcelWorksheet sourceWorksheet, ExcelRangeBase sourceRange)
    {
        string xml =
            "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"\" refreshOnLoad=\"1\" refreshedBy=\"SomeUser\" refreshedDate=\"40504.582403125001\" createdVersion=\"6\" refreshedVersion=\"6\" recordCount=\"5\" upgradeOnRefresh=\"1\">";

        xml += "<cacheSource type=\"worksheet\">";
        string? sourceName = sourceRange.GetName();

        if (string.IsNullOrEmpty(sourceName))
        {
            xml += string.Format("<worksheetSource ref=\"{0}\" sheet=\"{1}\" /> ", sourceRange.Address, sourceRange.WorkSheetName);
        }
        else
        {
            xml += string.Format("<worksheetSource name=\"{0}\" /> ", sourceName);
        }

        xml += "</cacheSource>";
        xml += string.Format("<cacheFields count=\"{0}\">", sourceRange._toCol - sourceRange._fromCol + 1);

        for (int col = sourceRange._fromCol; col <= sourceRange._toCol; col++)
        {
            object? name = sourceWorksheet?.GetValueInner(sourceRange._fromRow, col);

            if (name == null || name.ToString() == "")
            {
                xml += string.Format("<cacheField name=\"Column{0}\" numFmtId=\"0\">", col - sourceRange._fromCol + 1);
            }
            else
            {
                xml += string.Format("<cacheField name=\"{0}\" numFmtId=\"0\">", SecurityElement.Escape(name.ToString()));
            }

            xml += "<sharedItems containsBlank=\"1\" /> ";
            xml += "</cacheField>";
        }

        xml += "</cacheFields>";

        xml +=
            $"<extLst><ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{ExtLstUris.PivotCacheDefinitionUri}\"><x14:pivotCacheDefinition pivotCacheId=\"0\"/></ext></extLst>";

        xml += "</pivotCacheDefinition>";

        return xml;
    }

    internal void SetSourceName(string name)
    {
        this.DeleteNode(_sourceAddressPath); //Remove any address if previously set.
        this.SetXmlNodeString(_sourceNamePath, name);
    }

    internal void SetSourceAddress(string address)
    {
        this.DeleteNode(_sourceNamePath); //Remove any name or table if previously set.
        this.SetXmlNodeString(_sourceAddressPath, address);
    }

    int _cacheId = int.MinValue;

    internal int CacheId
    {
        get
        {
            if (this._cacheId < 0)
            {
                this._cacheId = this.GetXmlNodeInt("d:extLst/d:ext/x14:pivotCacheDefinition/@pivotCacheId");

                if (this._cacheId < 0)
                {
                    this._cacheId = this._wb.GetPivotCacheId(this.CacheDefinitionUri);
                    XmlNode? node = this.GetOrCreateExtLstSubNode(ExtLstUris.PivotCacheDefinitionUri, "x14");
                    node.InnerXml = $"<x14:pivotCacheDefinition pivotCacheId=\"{this._cacheId}\"/>";
                }
            }

            return this._cacheId;
        }
        set
        {
            XmlNode? node = this.GetOrCreateExtLstSubNode(ExtLstUris.PivotCacheDefinitionUri, "x14");

            if (node.InnerXml == "")
            {
                node.InnerXml = $"<x14:pivotCacheDefinition pivotCacheId=\"{this._cacheId}\"/>";
            }
            else
            {
                this.SetXmlNodeInt("d:extLst/d:ext/x14:pivotCacheDefinition/@pivotCacheId", value);
            }
        }
    }

    internal bool RefreshOnLoad
    {
        get { return this.GetXmlNodeBool("@refreshOnLoad"); }
        set { this.SetXmlNodeBool("@refreshOnLoad", value); }
    }

    public bool SaveData
    {
        get { return this.GetXmlNodeBool("@saveData", true); }
        set
        {
            if (this.SaveData == value)
            {
                return;
            }

            this.SetXmlNodeBool("@saveData", value);

            if (value)
            {
                this.AddRecordsXml();
            }
            else
            {
                this.RemoveRecordsXml();
            }

            this.SetXmlNodeBool("@saveData", value);
        }
    }

    private void RemoveRecordsXml()
    {
        this.RecordRelationshipId = null;
        this._wb._package.ZipPackage.DeletePart(this.CacheRecordUri);
        this.CacheRecordUri = null;
        this.RecordRelationship = null;
    }

    internal void AddRecordsXml()
    {
        int c = this.CacheId;

        //CacheRecord. Create an empty one.
        this.CacheRecordUri = GetNewUri(this._wb._package.ZipPackage, "/xl/pivotCache/pivotCacheRecords{0}.xml", ref c);
        this.ResetRecordXml(this._wb._package.ZipPackage);

        this.RecordRelationship = this.Part.CreateRelationship(UriHelper.ResolvePartUri(this.CacheDefinitionUri, this.CacheRecordUri),
                                                               TargetMode.Internal,
                                                               ExcelPackage.schemaRelationships + "/pivotCacheRecords");

        this.RecordRelationshipId = this.RecordRelationship.Id;
    }

    internal void Delete()
    {
        this._wb.RemovePivotTableCache(this.CacheId);
        this.Part.Package.DeletePart(this.CacheDefinitionUri);

        if (this.CacheRecordUri != null)
        {
            this.Part.Package.DeletePart(this.CacheRecordUri);
        }
    }

    internal ExcelPivotTableCacheField AddDateGroupField(ExcelPivotTableField field, eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int interval)
    {
        ExcelPivotTableCacheField cacheField = this.CreateField(groupBy.ToString(), field.Index, false);
        _ = cacheField.SetDateGroup(field, groupBy, startDate, endDate, interval);

        this.Fields.Add(cacheField);

        return cacheField;
    }

    internal ExcelPivotTableCacheField AddFormula(string name, string formula)
    {
        ExcelPivotTableCacheField cacheField = this.CreateField(name, this._fields.Count, false);
        cacheField.Formula = formula;
        this.Fields.Add(cacheField);

        return cacheField;
    }

    private ExcelPivotTableCacheField CreateField(string name, int index, bool databaseField = true)
    {
        //Add Cache definition field.
        XmlNode? cacheTopNode = this.CacheDefinitionXml.SelectSingleNode("//d:cacheFields", this.NameSpaceManager);
        XmlElement? cacheFieldNode = this.CacheDefinitionXml.CreateElement("cacheField", ExcelPackage.schemaMain);

        cacheFieldNode.SetAttribute("name", name);

        if (databaseField == false)
        {
            cacheFieldNode.SetAttribute("databaseField", "0");
        }

        _ = cacheTopNode.AppendChild(cacheFieldNode);

        return new ExcelPivotTableCacheField(this.NameSpaceManager, cacheFieldNode, this, index);
    }
}