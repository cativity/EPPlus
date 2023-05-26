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
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging;
using System.Linq;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Packaging.Ionic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Style.Dxf;
using System.IO;
using System.Globalization;
using OfficeOpenXml.Table.PivotTable.Filter;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Represents a null value in a pivot table caches shared items list.
    /// </summary>
    public struct PivotNull
    {
    }
    /// <summary>
    /// An Excel Pivottable
    /// </summary>
    public class ExcelPivotTable : XmlHelper
    {
        /// <summary>
        /// Represents a null value in a pivot table caches shared items list.
        /// </summary>
        public static PivotNull PivotNullValue = new PivotNull();
        internal ExcelPivotTable(ZipPackageRelationship rel, ExcelWorksheet sheet) :
            base(sheet.NameSpaceManager)
        {
            this.WorkSheet = sheet;
            this.PivotTableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            this.Relationship = rel;
            ZipPackage? pck = sheet._package.ZipPackage;
            this.Part = pck.GetPart(this.PivotTableUri);

            this.PivotTableXml = new XmlDocument();
            LoadXmlSafe(this.PivotTableXml, this.Part.GetStream());
            this.TopNode = this.PivotTableXml.DocumentElement;
            this.Init();
            this.Address = new ExcelAddressBase(this.GetXmlNodeString("d:location/@ref"));

            this.CacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this);
            this.LoadFields();

            int pos = 0;
            //Add row fields.
            foreach (XmlElement rowElem in this.TopNode.SelectNodes("d:rowFields/d:field", this.NameSpaceManager))
            {
                if (int.TryParse(rowElem.GetAttribute("x"), out int x) && x >= 0)
                {
                    this.RowFields.AddInternal(this.Fields[x]);
                }
                else
                {
                    if(x==-2)
                    {
                        this.ValuesFieldPosition = pos;
                    }
                    rowElem.ParentNode.RemoveChild(rowElem);
                }
                pos++;
            }

            pos = 0;
            ////Add column fields.
            foreach (XmlElement colElem in this.TopNode.SelectNodes("d:colFields/d:field", this.NameSpaceManager))
            {
                if (int.TryParse(colElem.GetAttribute("x"), out int x) && x >= 0)
                {
                    this.ColumnFields.AddInternal(this.Fields[x]);
                }
                else
                {
                    if (x == -2)
                    {
                        this.ValuesFieldPosition = pos;
                    }
                    colElem.ParentNode.RemoveChild(colElem);
                }
                pos++;
            }

            //Add Page elements
            //int index = 0;
            foreach (XmlElement pageElem in this.TopNode.SelectNodes("d:pageFields/d:pageField", this.NameSpaceManager))
            {
                if (int.TryParse(pageElem.GetAttribute("fld"), out int fld) && fld >= 0)
                {
                    ExcelPivotTableField? field = this.Fields[fld];
                    field._pageFieldSettings = new ExcelPivotTablePageFieldSettings(this.NameSpaceManager, pageElem, field, fld);
                    this.PageFields.AddInternal(field);
                }
            }

            //Add data elements
            //index = 0;
            foreach (XmlElement dataElem in this.TopNode.SelectNodes("d:dataFields/d:dataField", this.NameSpaceManager))
            {
                if (int.TryParse(dataElem.GetAttribute("fld"), out int fld) && fld >= 0)
                {
                    ExcelPivotTableField? field = this.Fields[fld];
                    ExcelPivotTableDataField? dataField = new ExcelPivotTableDataField(this.NameSpaceManager, dataElem, field);
                    this.DataFields.AddInternal(dataField);
                }
            }

            this.Styles = new ExcelPivotTableAreaStyleCollection(this);
        }
        /// <summary>
        /// Add a new pivottable
        /// </summary>
        /// <param name="sheet">The worksheet</param>
        /// <param name="address">the address of the pivottable</param>
        /// <param name="pivotTableCache">The pivot table cache</param>
        /// <param name="name"></param>
        /// <param name="tblId"></param>
        internal ExcelPivotTable(ExcelWorksheet sheet, ExcelAddressBase address, PivotTableCacheInternal pivotTableCache, string name, int tblId) :
        base(sheet.NameSpaceManager)
        {
            this.CreatePivotTable(sheet, address, pivotTableCache.Fields.Count, name, tblId);

            this.CacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, pivotTableCache);
            this.CacheId = pivotTableCache.CacheId;

            this.LoadFields();
            this.Styles = new ExcelPivotTableAreaStyleCollection(this);
        }
        /// <summary>
        /// Add a new pivottable
        /// </summary>
        /// <param name="sheet">The worksheet</param>
        /// <param name="address">the address of the pivottable</param>
        /// <param name="sourceAddress">The address of the Source data</param>
        /// <param name="name"></param>
        /// <param name="tblId"></param>
        internal ExcelPivotTable(ExcelWorksheet sheet, ExcelAddressBase address, ExcelRangeBase sourceAddress, string name, int tblId) :
        base(sheet.NameSpaceManager)
        {
            this.CreatePivotTable(sheet, address, sourceAddress._toCol - sourceAddress._fromCol + 1, name, tblId);

            this.CacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, sourceAddress);
            this.CacheId = this.CacheDefinition._cacheReference.CacheId;

            this.LoadFields();
            this.Styles = new ExcelPivotTableAreaStyleCollection(this);
        }

        private void CreatePivotTable(ExcelWorksheet sheet, ExcelAddressBase address, int fields, string name, int tblId)
        {
            this.WorkSheet = sheet;
            this.Address = address;
            ZipPackage? pck = sheet._package.ZipPackage;

            this.PivotTableXml = new XmlDocument();
            LoadXmlSafe(this.PivotTableXml, GetStartXml(name, address, fields), Encoding.UTF8);
            this.TopNode = this.PivotTableXml.DocumentElement;
            this.PivotTableUri = GetNewUri(pck, "/xl/pivotTables/pivotTable{0}.xml", ref tblId);
            this.Init();

            this.Part = pck.CreatePart(this.PivotTableUri, ContentTypes.contentTypePivotTable);
            this.PivotTableXml.Save(this.Part.GetStream());

            //Worksheet-Pivottable relationship
            this.Relationship = sheet.Part.CreateRelationship(UriHelper.ResolvePartUri(sheet.WorksheetUri, this.PivotTableUri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");

            using ExcelRange? r = sheet.Cells[address.Address];
            r.Clear();
        }

        private void Init()
        {
            this.SchemaNodeOrder = new string[] { "location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "dataFields", "formats", "conditionalFormats", "chartFormats", "pivotHierarchies", "pivotTableStyleInfo", "filters", "rowHierarchiesUsage", "colHierarchiesUsage", "extLst" };
        }
        private void LoadFields()
        {
            int index = 0;
            XmlNode? pivotFieldNode = this.TopNode.SelectSingleNode("d:pivotFields", this.NameSpaceManager);
            //Add fields.            
            foreach (XmlElement fieldElem in pivotFieldNode.SelectNodes("d:pivotField", this.NameSpaceManager))
            {
                ExcelPivotTableField? fld = new ExcelPivotTableField(this.NameSpaceManager, fieldElem, this, index, index);
                fld._cacheField = this.CacheDefinition._cacheReference.Fields[index++];
                fld.LoadItems();
                this.Fields.AddInternal(fld);
            }

        }
        private static string GetStartXml(string name, ExcelAddressBase address, int fields)
        {
            string xml = string.Format("<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"{0}\" dataOnRows=\"1\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" dataCaption=\"Data\"  createdVersion=\"6\" updatedVersion=\"6\" showMemberPropertyTips=\"0\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" indent=\"0\" compact=\"0\" compactData=\"0\" gridDropZones=\"1\">",
                ConvertUtil.ExcelEscapeString(name));

            xml += string.Format("<location ref=\"{0}\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" /> ", address.FirstAddress);
            xml += string.Format("<pivotFields count=\"{0}\">", fields);
            for (int col = 0; col < fields; col++)
            {
                xml += "<pivotField showAll=\"0\" />"; //compact=\"0\" outline=\"0\" subtotalTop=\"0\" includeNewItemsInFilter=\"1\"     
            }

            xml += "</pivotFields>";
            xml += "<pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\" />";
            xml += $"<extLst><ext xmlns:xpdl=\"http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout\" uri=\"{ExtLstUris.PivotTableDefinition16Uri }\"><xpdl:pivotTableDefinition16/></ext></extLst>";
            xml += "</pivotTableDefinition>";
            return xml;
        }
        internal ZipPackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Individual styles for the pivot table.
        /// </summary>
        public ExcelPivotTableAreaStyleCollection Styles
        {
            get;
            private set;
        }
        /// <summary>
        /// Provides access to the XML data representing the pivottable in the package.
        /// </summary>
        public XmlDocument PivotTableXml { get; private set; }
        /// <summary>
        /// The package internal URI to the pivottable Xml Document.
        /// </summary>
        public Uri PivotTableUri
        {
            get;
            internal set;
        }
        internal ZipPackageRelationship Relationship
        {
            get;
            set;
        }
        const string NAME_PATH = "@name";
        const string DISPLAY_NAME_PATH = "@displayName";
        /// <summary>
        /// Name of the pivottable object in Excel
        /// </summary>
        public string Name
        {
            get
            {
                return this.GetXmlNodeString(NAME_PATH);
            }
            set
            {
                if (this.WorkSheet.Workbook.ExistsTableName(value))
                {
                    throw (new ArgumentException("PivotTable name is not unique"));
                }
                string prevName = this.Name;
                if (this.WorkSheet.Tables._tableNames.ContainsKey(prevName))
                {
                    int ix = this.WorkSheet.Tables._tableNames[prevName];
                    this.WorkSheet.Tables._tableNames.Remove(prevName);
                    this.WorkSheet.Tables._tableNames.Add(value, ix);
                }

                this.SetXmlNodeString(NAME_PATH, value);
                this.SetXmlNodeString(DISPLAY_NAME_PATH, cleanDisplayName(value));
            }
        }
        /// <summary>
        /// Reference to the pivot table cache definition object
        /// </summary>
        public ExcelPivotCacheDefinition CacheDefinition
        {
            get;
            private set;
        }
        private static string cleanDisplayName(string name)
        {
            return Regex.Replace(name, @"[^\w\.-_]", "_");
        }
        #region "Public Properties"

        /// <summary>
        /// The worksheet where the pivottable is located
        /// </summary>
        public ExcelWorksheet WorkSheet
        {
            get;
            set;
        }
        /// <summary>
        /// The location of the pivot table
        /// </summary>
        public ExcelAddressBase Address
        {
            get;
            internal set;
        }
        /// <summary>
        /// If multiple datafields are displayed in the row area or the column area
        /// </summary>
        public bool DataOnRows
        {
            get
            {
                return this.GetXmlNodeBool("@dataOnRows");
            }
            set
            {
                this.SetXmlNodeBool("@dataOnRows", value);
            }
        }
        /// <summary>
        /// The position of the values in the row- or column- fields list. Position is dependent on <see cref="DataOnRows"/>.
        /// If DataOnRows is true then the position is within the <see cref="ColumnFields"/> collection,
        /// a value of false the position is within the <see cref="RowFields" /> collection.
        /// A negative value or a value out of range of the add the "Σ values" field to the end of the collection.
        /// </summary>
        public int ValuesFieldPosition
        {
            get;
            set;
        } = -1;
        /// <summary>
        /// if true apply legacy table autoformat number format properties.
        /// </summary>
        public bool ApplyNumberFormats
        {
            get
            {
                return this.GetXmlNodeBool("@applyNumberFormats");
            }
            set
            {
                this.SetXmlNodeBool("@applyNumberFormats", value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat border properties
        /// </summary>
        public bool ApplyBorderFormats
        {
            get
            {
                return this.GetXmlNodeBool("@applyBorderFormats");
            }
            set
            {
                this.SetXmlNodeBool("@applyBorderFormats", value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat font properties
        /// </summary>
        public bool ApplyFontFormats
        {
            get
            {
                return this.GetXmlNodeBool("@applyFontFormats");
            }
            set
            {
                this.SetXmlNodeBool("@applyFontFormats", value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat pattern properties
        /// </summary>
        public bool ApplyPatternFormats
        {
            get
            {
                return this.GetXmlNodeBool("@applyPatternFormats");
            }
            set
            {
                this.SetXmlNodeBool("@applyPatternFormats", value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat width/height properties.
        /// </summary>
        public bool ApplyWidthHeightFormats
        {
            get
            {
                return this.GetXmlNodeBool("@applyWidthHeightFormats");
            }
            set
            {
                this.SetXmlNodeBool("@applyWidthHeightFormats", value);
            }
        }
        /// <summary>
        /// Show member property information
        /// </summary>
        public bool ShowMemberPropertyTips
        {
            get
            {
                return this.GetXmlNodeBool("@showMemberPropertyTips");
            }
            set
            {
                this.SetXmlNodeBool("@showMemberPropertyTips", value);
            }
        }
        /// <summary>
        /// Show the drill indicators
        /// </summary>
        public bool ShowCalcMember
        {
            get
            {
                return this.GetXmlNodeBool("@showCalcMbrs");
            }
            set
            {
                this.SetXmlNodeBool("@showCalcMbrs", value);
            }
        }
        /// <summary>
        /// If the user is prevented from drilling down on a PivotItem or aggregate value
        /// </summary>
        public bool EnableDrill
        {
            get
            {
                return this.GetXmlNodeBool("@enableDrill", true);
            }
            set
            {
                this.SetXmlNodeBool("@enableDrill", value);
            }
        }
        /// <summary>
        /// Show the drill down buttons
        /// </summary>
        public bool ShowDrill
        {
            get
            {
                return this.GetXmlNodeBool("@showDrill", true);
            }
            set
            {
                this.SetXmlNodeBool("@showDrill", value);
            }
        }
        /// <summary>
        /// If the tooltips should be displayed for PivotTable data cells.
        /// </summary>
        public bool ShowDataTips
        {
            get
            {
                return this.GetXmlNodeBool("@showDataTips", true);
            }
            set
            {
                this.SetXmlNodeBool("@showDataTips", value, true);
            }
        }
        /// <summary>
        /// If the row and column titles from the PivotTable should be printed.
        /// </summary>
        public bool FieldPrintTitles
        {
            get
            {
                return this.GetXmlNodeBool("@fieldPrintTitles");
            }
            set
            {
                this.SetXmlNodeBool("@fieldPrintTitles", value);
            }
        }
        /// <summary>
        /// If the row and column titles from the PivotTable should be printed.
        /// </summary>
        public bool ItemPrintTitles
        {
            get
            {
                return this.GetXmlNodeBool("@itemPrintTitles");
            }
            set
            {
                this.SetXmlNodeBool("@itemPrintTitles", value);
            }
        }
        /// <summary>
        /// If the grand totals should be displayed for the PivotTable columns
        /// </summary>
        public bool ColumnGrandTotals
        {
            get
            {
                return this.GetXmlNodeBool("@colGrandTotals");
            }
            set
            {
                this.SetXmlNodeBool("@colGrandTotals", value);
            }
        }
        /// <summary>
        /// If the grand totals should be displayed for the PivotTable rows
        /// </summary>
        public bool RowGrandTotals
        {
            get
            {
                return this.GetXmlNodeBool("@rowGrandTotals");
            }
            set
            {
                this.SetXmlNodeBool("@rowGrandTotals", value);
            }
        }
        /// <summary>
        /// If the drill indicators expand collapse buttons should be printed.
        /// </summary>
        public bool PrintDrill
        {
            get
            {
                return this.GetXmlNodeBool("@printDrill");
            }
            set
            {
                this.SetXmlNodeBool("@printDrill", value);
            }
        }
        /// <summary>
        /// Indicates whether to show error messages in cells.
        /// </summary>
        public bool ShowError
        {
            get
            {
                return this.GetXmlNodeBool("@showError");
            }
            set
            {
                this.SetXmlNodeBool("@showError", value);
            }
        }
        /// <summary>
        /// The string to be displayed in cells that contain errors.
        /// </summary>
        public string ErrorCaption
        {
            get
            {
                return this.GetXmlNodeString("@errorCaption");
            }
            set
            {
                this.SetXmlNodeString("@errorCaption", value);
            }
        }
        /// <summary>
        /// Specifies the name of the value area field header in the PivotTable. 
        /// This caption is shown when the PivotTable when two or more fields are in the values area.
        /// </summary>
        public string DataCaption
        {
            get
            {
                return this.GetXmlNodeString("@dataCaption");
            }
            set
            {
                this.SetXmlNodeString("@dataCaption", value);
            }
        }
        /// <summary>
        /// Show field headers
        /// </summary>
        public bool ShowHeaders
        {
            get
            {
                return this.GetXmlNodeBool("@showHeaders");
            }
            set
            {
                this.SetXmlNodeBool("@showHeaders", value);
            }
        }
        /// <summary>
        /// The number of page fields to display before starting another row or column
        /// </summary>
        public int PageWrap
        {
            get
            {
                return this.GetXmlNodeInt("@pageWrap");
            }
            set
            {
                if (value < 0)
                {
                    throw new Exception("Value can't be negative");
                }

                this.SetXmlNodeString("@pageWrap", value.ToString());
            }
        }
        /// <summary>
        /// A boolean that indicates whether legacy auto formatting has been applied to the PivotTable view
        /// </summary>
        public bool UseAutoFormatting
        {
            get
            {
                return this.GetXmlNodeBool("@useAutoFormatting");
            }
            set
            {
                this.SetXmlNodeBool("@useAutoFormatting", value);
            }
        }
        /// <summary>
        /// A boolean that indicates if the in-grid drop zones should be displayed at runtime, and if classic layout is applied
        /// </summary>
        public bool GridDropZones
        {
            get
            {
                return this.GetXmlNodeBool("@gridDropZones");
            }
            set
            {
                this.SetXmlNodeBool("@gridDropZones", value);
            }
        }
        /// <summary>
        /// The indentation increment for compact axis and can be used to set the Report Layout to Compact Form
        /// </summary>
        public int Indent
        {
            get
            {
                return this.GetXmlNodeInt("@indent");
            }
            set
            {
                this.SetXmlNodeString("@indent", value.ToString());
            }
        }
        /// <summary>
        /// A boolean that indicates whether data fields in the PivotTable should be displayed in outline form
        /// </summary>
        public bool OutlineData
        {
            get
            {
                return this.GetXmlNodeBool("@outlineData");
            }
            set
            {
                this.SetXmlNodeBool("@outlineData", value);
            }
        }
        /// <summary>
        /// A boolean that indicates whether new fields should have their outline flag set to true
        /// </summary>
        public bool Outline
        {
            get
            {
                return this.GetXmlNodeBool("@outline");
            }
            set
            {
                this.SetXmlNodeBool("@outline", value);
            }
        }
        /// <summary>
        /// A boolean that indicates if the fields of a PivotTable can have multiple filters set on them
        /// </summary>
        public bool MultipleFieldFilters
        {
            get
            {
                return this.GetXmlNodeBool("@multipleFieldFilters");
            }
            set
            {
                this.SetXmlNodeBool("@multipleFieldFilters", value);
            }
        }
        /// <summary>
        /// A boolean that indicates if new fields should have their compact flag set to true
        /// </summary>
        public bool Compact
        {
            get
            {
                return this.GetXmlNodeBool("@compact");
            }
            set
            {
                this.SetXmlNodeBool("@compact", value);
            }
        }
        /// <summary>
        /// Sets all pivot table fields <see cref="ExcelPivotTableField.Compact"/> property to the value supplied.
        /// </summary>
        /// <param name="value">The the value for the Compact property.</param>
        public void SetCompact(bool value=true)
        {
            this.Compact = value;
            foreach(ExcelPivotTableField? f in this.Fields)
            {
                f.Compact = value;
            }
        }
        /// <summary>
        /// A boolean that indicates if the field next to the data field in the PivotTable should be displayed in the same column of the spreadsheet.
        /// </summary>
        public bool CompactData
        {
            get
            {
                return this.GetXmlNodeBool("@compactData");
            }
            set
            {
                this.SetXmlNodeBool("@compactData", value);
            }
        }
        /// <summary>
        /// Specifies the string to be displayed for grand totals.
        /// </summary>
        public string GrandTotalCaption
        {
            get
            {
                return this.GetXmlNodeString("@grandTotalCaption");
            }
            set
            {
                this.SetXmlNodeString("@grandTotalCaption", value);
            }
        }
        /// <summary>
        /// The text to be displayed in row header in compact mode.
        /// </summary>
        public string RowHeaderCaption
        {
            get
            {
                return this.GetXmlNodeString("@rowHeaderCaption");
            }
            set
            {
                this.SetXmlNodeString("@rowHeaderCaption", value);
            }
        }
        /// <summary>
        /// The text to be displayed in column header in compact mode.
        /// </summary>
        public string ColumnHeaderCaption
        {
            get
            {
                return this.GetXmlNodeString("@colHeaderCaption");
            }
            set
            {
                this.SetXmlNodeString("@colHeaderCaption", value);
            }
        }
        /// <summary>
        /// The text to be displayed in cells with no value
        /// </summary>
        public string MissingCaption
        {
            get
            {
                return this.GetXmlNodeString("@missingCaption");
            }
            set
            {
                this.SetXmlNodeString("@missingCaption", value);
            }
        }
        ExcelPivotTableFilterCollection _filters = null;
        /// <summary>
        /// Filters applied to the pivot table
        /// </summary>
        public ExcelPivotTableFilterCollection Filters
        {
            get
            {
                if (this._filters == null)
                {
                    this._filters = new ExcelPivotTableFilterCollection(this);
                }
                return this._filters;
            }
        }
        const string FIRSTHEADERROW_PATH = "d:location/@firstHeaderRow";
        /// <summary>
        /// The first row of the PivotTable header, relative to the top left cell in the ref value
        /// </summary>
        public int FirstHeaderRow
        {
            get
            {
                return this.GetXmlNodeInt(FIRSTHEADERROW_PATH);
            }
            set
            {
                this.SetXmlNodeString(FIRSTHEADERROW_PATH, value.ToString());
            }
        }
        const string FIRSTDATAROW_PATH = "d:location/@firstDataRow";
        /// <summary>
        /// The first column of the PivotTable data, relative to the top left cell in the range
        /// </summary>
        public int FirstDataRow
        {
            get
            {
                return this.GetXmlNodeInt(FIRSTDATAROW_PATH);
            }
            set
            {
                this.SetXmlNodeString(FIRSTDATAROW_PATH, value.ToString());
            }
        }
        const string FIRSTDATACOL_PATH = "d:location/@firstDataCol";
        /// <summary>
        /// The first column of the PivotTable data, relative to the top left cell in the range.
        /// </summary>
        public int FirstDataCol
        {
            get
            {
                return this.GetXmlNodeInt(FIRSTDATACOL_PATH);
            }
            set
            {
                this.SetXmlNodeString(FIRSTDATACOL_PATH, value.ToString());
            }
        }
        ExcelPivotTableFieldCollection _fields = null;
        /// <summary>
        /// The fields in the table 
        /// </summary>
        public ExcelPivotTableFieldCollection Fields
        {
            get
            {
                if (this._fields == null)
                {
                    this._fields = new ExcelPivotTableFieldCollection(this);
                }
                return this._fields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _rowFields = null;
        /// <summary>
        /// Row label fields 
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection RowFields
        {
            get
            {
                if (this._rowFields == null)
                {
                    this._rowFields = new ExcelPivotTableRowColumnFieldCollection(this, "rowFields");
                }
                return this._rowFields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _columnFields = null;
        /// <summary>
        /// Column label fields 
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection ColumnFields
        {
            get
            {
                if (this._columnFields == null)
                {
                    this._columnFields = new ExcelPivotTableRowColumnFieldCollection(this, "colFields");
                }
                return this._columnFields;
            }
        }
        ExcelPivotTableDataFieldCollection _dataFields = null;
        /// <summary>
        /// Value fields 
        /// </summary>
        public ExcelPivotTableDataFieldCollection DataFields
        {
            get
            {
                if (this._dataFields == null)
                {
                    this._dataFields = new ExcelPivotTableDataFieldCollection(this);
                }
                return this._dataFields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _pageFields = null;
        /// <summary>
        /// Report filter fields
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection PageFields
        {
            get
            {
                if (this._pageFields == null)
                {
                    this._pageFields = new ExcelPivotTableRowColumnFieldCollection(this, "pageFields");
                }
                return this._pageFields;
            }
        }
        const string STYLENAME_PATH = "d:pivotTableStyleInfo/@name";
        /// <summary>
        /// Pivot style name. Used for custom styles
        /// </summary>
        public string StyleName
        {
            get
            {
                return this.GetXmlNodeString(STYLENAME_PATH);
            }
            set
            {
                if (value.StartsWith("PivotStyle"))
                {
                    try
                    {
                        if (Enum.GetNames(typeof(TableStyles)).Any(x => x.Equals(value.Substring(10, value.Length - 10), StringComparison.OrdinalIgnoreCase)))
                        {
                            this._tableStyle = (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10, value.Length - 10), true);
                        }
                        else
                        {
                            this._tableStyle = TableStyles.Custom;
                        }
                    }
                    catch
                    {
                        this._tableStyle = TableStyles.Custom;
                    }
                    try
                    {
                        this._pivotTableStyle = (PivotTableStyles)Enum.Parse(typeof(PivotTableStyles), value.Substring(10, value.Length - 10), true);
                    }
                    catch
                    {
                        this._pivotTableStyle = PivotTableStyles.Custom;
                    }

                }
                else if (value == "None")
                {
                    this._tableStyle = TableStyles.None;
                    this._pivotTableStyle = PivotTableStyles.None;
                    value = "";
                }
                else
                {
                    this._tableStyle = TableStyles.Custom;
                    this._pivotTableStyle = PivotTableStyles.Custom;
                }

                this.SetXmlNodeString(STYLENAME_PATH, value, true);
            }
        }
        const string SHOWCOLHEADERS_PATH = "d:pivotTableStyleInfo/@showColHeaders";
        /// <summary>
        /// Whether to show column headers for the pivot table.
        /// </summary>
        public bool ShowColumnHeaders
        {
            get
            {
                return this.GetXmlNodeBool(SHOWCOLHEADERS_PATH);
            }
            set
            {
                this.SetXmlNodeBool(SHOWCOLHEADERS_PATH, value);
            }
        }
        const string SHOWCOLSTRIPES_PATH = "d:pivotTableStyleInfo/@showColStripes";
        /// <summary>
        /// Whether to show column stripe formatting for the pivot table.
        /// </summary>
        public bool ShowColumnStripes
        {
            get
            {
                return this.GetXmlNodeBool(SHOWCOLSTRIPES_PATH);
            }
            set
            {
                this.SetXmlNodeBool(SHOWCOLSTRIPES_PATH, value);
            }
        }
        const string SHOWLASTCOLUMN_PATH = "d:pivotTableStyleInfo/@showLastColumn";
        /// <summary>
        /// Whether to show the last column for the pivot table.
        /// </summary>
        public bool ShowLastColumn
        {
            get
            {
                return this.GetXmlNodeBool(SHOWLASTCOLUMN_PATH);
            }
            set
            {
                this.SetXmlNodeBool(SHOWLASTCOLUMN_PATH, value);
            }
        }
        const string SHOWROWHEADERS_PATH = "d:pivotTableStyleInfo/@showRowHeaders";
        /// <summary>
        /// Whether to show row headers for the pivot table.
        /// </summary>
        public bool ShowRowHeaders
        {
            get
            {
                return this.GetXmlNodeBool(SHOWROWHEADERS_PATH);
            }
            set
            {
                this.SetXmlNodeBool(SHOWROWHEADERS_PATH, value);
            }
        }
        const string SHOWROWSTRIPES_PATH = "d:pivotTableStyleInfo/@showRowStripes";
        /// <summary>
        /// Whether to show row stripe formatting for the pivot table.
        /// </summary>
        public bool ShowRowStripes
        {
            get
            {
                return this.GetXmlNodeBool(SHOWROWSTRIPES_PATH);
            }
            set
            {
                this.SetXmlNodeBool(SHOWROWSTRIPES_PATH, value);
            }
        }
        TableStyles _tableStyle = TableStyles.Medium6;
        /// <summary>
        /// The table style. If this property is Custom, the style from the StyleName propery is used.
        /// </summary>
        [Obsolete("Use the PivotTableStyle property for more options")]
        public TableStyles TableStyle
        {
            get
            {
                return this._tableStyle;
            }
            set
            {
                this._tableStyle = value;
                if (value != TableStyles.Custom)
                {
                    this.StyleName = "PivotStyle" + value.ToString();
                }
            }
        }
        PivotTableStyles _pivotTableStyle = PivotTableStyles.Medium6;
        /// <summary>
        /// The pivot table style. If this property is Custom, the style from the StyleName propery is used.
        /// </summary>
        public PivotTableStyles PivotTableStyle
        {
            get
            {
                return this._pivotTableStyle;
            }
            set
            {
                this._pivotTableStyle = value;
                if (value != PivotTableStyles.Custom)
                {
                    //SetXmlNodeString(STYLENAME_PATH, "PivotStyle" + value.ToString());
                    this.StyleName = "PivotStyle" + value.ToString();
                }
            }
        }
        const string _showValuesRowPath = "d:extLst/d:ext[@uri='" + ExtLstUris.PivotTableDefinitionUri + "']/x14:pivotTableDefinition/@hideValuesRow";
        /// <summary>
        /// If the pivot tables value row is visible or not. 
        /// This property only applies when <see cref="GridDropZones"/> is set to false.
        /// </summary>
        public bool ShowValuesRow
        {
            get
            {
                return !this.GetXmlNodeBool(_showValuesRowPath);
            }
            set
            {
                XmlNode? node = this.GetOrCreateExtLstSubNode(ExtLstUris.PivotTableDefinitionUri, "x14");
                XmlHelper? xh = XmlHelperFactory.Create(this.NameSpaceManager, node);
                xh.SetXmlNodeBool("x14:pivotTableDefinition/@hideValuesRow", !value);
            }
        }

        #endregion
        #region "Internal Properties"
        internal int CacheId
        {
            get
            {
                return this.GetXmlNodeInt("@cacheId", 0);
            }
            set
            {
                this.SetXmlNodeInt("@cacheId", value);
            }
        }

        internal int ChangeCacheId(int oldCacheId)
        {
            int newCacheId = this.WorkSheet.Workbook.GetNewPivotCacheId();
            this.CacheId = newCacheId;
            this.CacheDefinition._cacheReference.CacheId = newCacheId;
            this.WorkSheet.Workbook.SetXmlNodeInt($"d:pivotCaches/d:pivotCache[@cacheId={oldCacheId}]/@cacheId", newCacheId);

            return newCacheId;
        }

        #endregion
        int _newFilterId = 0;
        internal int GetNewFilterId()
        {
            return this._newFilterId++;
        }
        internal void SetNewFilterId(int value)
        {
            if (value >= this._newFilterId)
            {
                this._newFilterId = value + 1;
            }
        }

        internal void Save()
        {
            if(this.CacheDefinition.CacheSource==eSourceType.Worksheet)
            {
                if(this.CacheDefinition.SourceRange.Columns!= this.Fields.Count)
                {   
                    //if(Fields.Count)
                    //CacheDefinition.Refresh();
                }
            }
            if (this.DataFields.Count > 1)
            {
                XmlElement parentNode;
                int fields;
                if (this.DataOnRows == true)
                {
                    parentNode = this.PivotTableXml.SelectSingleNode("//d:rowFields", this.NameSpaceManager) as XmlElement;
                    if (parentNode == null)
                    {
                        this.CreateNode("d:rowFields");
                        parentNode = this.PivotTableXml.SelectSingleNode("//d:rowFields", this.NameSpaceManager) as XmlElement;
                    }
                    fields = this.RowFields.Count;
                }
                else
                {
                    parentNode = this.PivotTableXml.SelectSingleNode("//d:colFields", this.NameSpaceManager) as XmlElement;
                    if (parentNode == null)
                    {
                        this.CreateNode("d:colFields");
                        parentNode = this.PivotTableXml.SelectSingleNode("//d:colFields", this.NameSpaceManager) as XmlElement;
                    }
                    fields = this.ColumnFields.Count;
                }

                if (parentNode.SelectSingleNode("d:field[@ x= \"-2\"]", this.NameSpaceManager) == null)
                {
                    XmlElement fieldNode = this.PivotTableXml.CreateElement("field", ExcelPackage.schemaMain);
                    fieldNode.SetAttribute("x", "-2");
                    if (this.ValuesFieldPosition >= 0 && this.ValuesFieldPosition < fields)
                    {
                        parentNode.InsertBefore(fieldNode, parentNode.ChildNodes[this.ValuesFieldPosition]);
                    }
                    else
                    {
                        parentNode.AppendChild(fieldNode);
                    }
                }
            }

            this.SetXmlNodeString("d:location/@ref", this.Address.Address);

            foreach (ExcelPivotTableField? field in this.Fields)
            {
                field.SaveToXml();
            }

            foreach (ExcelPivotTableDataField? df in this.DataFields)
            {
                if (string.IsNullOrEmpty(df.Name))
                {

                    string name;
                    if (df.Function == DataFieldFunctions.None)
                    {
                        name = df.Field.Name; //Name must be set or Excel will crash on rename.                                
                    }
                    else
                    {
                        name = df.Function.ToString() + " of " + df.Field.Name; //Name must be set or Excel will crash on rename.
                    }

                    //Make sure name is unique
                    string? newName = name;
                    int i = 2;
                    while (this.DataFields.ExistsDfName(newName, df))
                    {
                        newName = name + (i++).ToString(CultureInfo.InvariantCulture);
                    }
                    df.Name = newName;
                }
            }

            this.UpdatePivotTableStyles();
            this.PivotTableXml.Save(this.Part.GetStream(FileMode.Create));
        }

        private void UpdatePivotTableStyles()
        {
            foreach (ExcelPivotTableAreaStyle a in this.Styles)
            {
                a.Conditions.UpdateXml();
            }
        }
    }
}
