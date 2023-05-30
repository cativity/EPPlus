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
using OfficeOpenXml.Utils;
using System.Security;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Core.Worksheet;
using System.Data;
using OfficeOpenXml.Export.ToDataTable;
using System.IO;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Export.HtmlExport;
using System.Globalization;
using OfficeOpenXml.Sorting;
using OfficeOpenXml.Export.HtmlExport.Interfaces;
using OfficeOpenXml.Packaging;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// An Excel Table
    /// </summary>
    public class ExcelTable : ExcelTableDxfBase, IEqualityComparer<ExcelTable>
    {
        internal ExcelTable(ZipPackageRelationship rel, ExcelWorksheet sheet)
            : base(sheet.NameSpaceManager)
        {
            this.WorkSheet = sheet;
            this.TableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            this.RelationshipID = rel.Id;
            ZipPackage? pck = sheet._package.ZipPackage;
            this.Part = pck.GetPart(this.TableUri);

            this.TableXml = new XmlDocument();
            LoadXmlSafe(this.TableXml, this.Part.GetStream());
            this.Init();
            this.Address = new ExcelAddressBase(this.GetXmlNodeString("@ref"));
            this._tableStyle = GetTableStyle(this.StyleName);
        }

        internal ExcelTable(ExcelWorksheet sheet, ExcelAddressBase address, string name, int tblId)
            : base(sheet.NameSpaceManager)
        {
            this.WorkSheet = sheet;
            this._address = address;

            this.TableXml = new XmlDocument();
            LoadXmlSafe(this.TableXml, this.GetStartXml(name, tblId), Encoding.UTF8);

            this.Init();

            //If the table is just one row we cannot have a header.
            if (address._fromRow == address._toRow)
            {
                this.ShowHeader = false;
            }

            if (this.AutoFilterAddress != null)
            {
                this.SetAutoFilter();
            }
        }

        private void Init()
        {
            this.TopNode = this.TableXml.DocumentElement;
            this.SchemaNodeOrder = new string[] { "autoFilter", "sortState", "tableColumns", "tableStyleInfo" };
            this.InitDxf(this.WorkSheet.Workbook.Styles, this, null);
            this.TableBorderStyle = new ExcelDxfBorderBase(this.WorkSheet.Workbook.Styles, null);
            this.HeaderRowBorderStyle = new ExcelDxfBorderBase(this.WorkSheet.Workbook.Styles, null);
            this._tableSorter = new TableSorter(this);
        }

        private string GetStartXml(string name, int tblId)
        {
            name = ConvertUtil.ExcelEscapeString(name);
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>";

            xml +=
                string.Format("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{0}\" name=\"{1}\" displayName=\"{2}\" ref=\"{3}\" headerRowCount=\"1\">",
                              tblId,
                              name,
                              ExcelAddressUtil.GetValidName(name),
                              this.Address.Address);

            xml += string.Format("<autoFilter ref=\"{0}\" />", this.Address.Address);

            int cols = this.Address._toCol - this.Address._fromCol + 1;
            xml += string.Format("<tableColumns count=\"{0}\">", cols);
            HashSet<string>? names = new HashSet<string>();

            for (int i = 1; i <= cols; i++)
            {
                ExcelRange? cell = this.WorkSheet.Cells[this.Address._fromRow, this.Address._fromCol + i - 1];
                string colName = SecurityElement.Escape(cell.Value?.ToString());

                if (cell.Value == null || names.Contains(colName))
                {
                    //Get an unique name
                    int a = i;

                    do
                    {
                        colName = string.Format("Column{0}", a++);
                    } while (names.Contains(colName));
                }

                _ = names.Add(colName);
                xml += string.Format("<tableColumn id=\"{0}\" name=\"{1}\" />", i, colName);
            }

            xml += "</tableColumns>";
            xml += "<tableStyleInfo name=\"TableStyleMedium9\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\" /> ";
            xml += "</table>";

            return xml;
        }

        internal static string CleanDisplayName(string name)
        {
            return Regex.Replace(name, @"[^\w\.-_]", "_");
        }

        internal ZipPackagePart Part { get; set; }

        /// <summary>
        /// Provides access to the XML data representing the table in the package.
        /// </summary>
        public XmlDocument TableXml { get; set; }

        /// <summary>
        /// The package internal URI to the Table Xml Document.
        /// </summary>
        public Uri TableUri { get; internal set; }

        internal string RelationshipID { get; set; }

        const string ID_PATH = "@id";

        internal int Id
        {
            get { return this.GetXmlNodeInt(ID_PATH); }
            set { this.SetXmlNodeString(ID_PATH, value.ToString()); }
        }

        const string NAME_PATH = "@name";
        const string DISPLAY_NAME_PATH = "@displayName";

        /// <summary>
        /// The name of the table object in Excel
        /// </summary>
        public string Name
        {
            get { return this.GetXmlNodeString(NAME_PATH); }
            set
            {
                if (this.Name.Equals(value, StringComparison.CurrentCultureIgnoreCase) == false && this.WorkSheet.Workbook.ExistsTableName(value))
                {
                    throw new ArgumentException("Tablename is not unique");
                }

                string prevName = this.Name;

                if (this.WorkSheet.Tables._tableNames.ContainsKey(prevName))
                {
                    int ix = this.WorkSheet.Tables._tableNames[prevName];
                    _ = this.WorkSheet.Tables._tableNames.Remove(prevName);
                    this.WorkSheet.Tables._tableNames.Add(value, ix);
                }

                TableAdjustFormula? ta = new TableAdjustFormula(this);
                ta.AdjustFormulas(prevName, value);
                this.SetXmlNodeString(NAME_PATH, value);
                this.SetXmlNodeString(DISPLAY_NAME_PATH, ExcelAddressUtil.GetValidName(value));
            }
        }

        internal void DeleteMe()
        {
            if (this.RelationshipID != null)
            {
                this.WorkSheet.DeleteNode($"d:tableParts/d:tablePart[@r:id='{this.RelationshipID}']");
            }

            if (this.TableUri != null && this.WorkSheet._package.ZipPackage.PartExists(this.TableUri))
            {
                this.WorkSheet._package.ZipPackage.DeletePart(this.TableUri);
            }
        }

        /// <summary>
        /// The worksheet of the table
        /// </summary>
        public ExcelWorksheet WorkSheet { get; set; }

        private ExcelAddressBase _address;

        /// <summary>
        /// The address of the table
        /// </summary>
        public ExcelAddressBase Address
        {
            get { return this._address; }
            internal set
            {
                this._address = value;

                if (value != null)
                {
                    this.SetXmlNodeString("@ref", value.Address);
                    this.WriteAutoFilter(this.ShowTotal);
                }
            }
        }

        /// <summary>
        /// The table range
        /// </summary>
        public ExcelRangeBase Range
        {
            get { return this.WorkSheet.Cells[this._address._fromRow, this._address._fromCol, this._address._toRow, this._address._toCol]; }
        }

        internal ExcelRangeBase DataRange
        {
            get
            {
                int fromRow = this.ShowHeader ? this._address._fromRow + 1 : this._address._fromRow;
                int toRow = this.ShowTotal ? this._address._toRow - 1 : this._address._toRow;

                return this.WorkSheet.Cells[fromRow, this._address._fromCol, toRow, this._address._toCol];
            }
        }

        #region Export table data

        /// <summary>
        /// Converts the table range to CSV format
        /// </summary>
        /// <returns></returns>
        /// <seealso cref="ExcelRangeBase.ToText()"/>
        public string ToText()
        {
            return this.Range.ToText();
        }

        /// <summary>
        /// Creates an <see cref="IExcelHtmlTableExporter"/> object to export the table to HTML
        /// </summary>
        /// <returns>The exporter object</returns>
        public IExcelHtmlTableExporter CreateHtmlExporter()
        {
            return new Export.HtmlExport.Exporters.ExcelHtmlTableExporter(this);
        }

        /// <summary>
        /// Converts the table range to CSV format
        /// </summary>
        /// <param name="format">Parameters/options for conversion to text</param>
        /// <returns></returns>
        /// <seealso cref="ExcelRangeBase.ToText(ExcelOutputTextFormat)"/>
        public string ToText(ExcelOutputTextFormat format)
        {
            return this.Range.ToText(format);
        }

#if !NET35 && !NET40
        /// <summary>
        /// Converts the table range to CSV format
        /// </summary>
        /// <returns></returns>
        /// <seealso cref="ExcelRangeBase.ToTextAsync()"/>
        public Task<string> ToTextAsync()
        {
            return this.Range.ToTextAsync();
        }

        /// <summary>
        /// Converts the table range to CSV format
        /// </summary>
        /// <returns></returns>
        /// <seealso cref="ExcelRangeBase.ToText(ExcelOutputTextFormat)"/>
        public Task<string> ToTextAsync(ExcelOutputTextFormat format)
        {
            return this.Range.ToTextAsync(format);
        }
#endif

        /// <summary>
        /// Exports the table to a file
        /// </summary>
        /// <param name="file">The export file</param>
        /// <param name="format">Export options</param>
        /// <seealso cref="ExcelRangeBase.SaveToText(FileInfo, ExcelOutputTextFormat)"></seealso>
        public void SaveToText(FileInfo file, ExcelOutputTextFormat format)
        {
            this.Range.SaveToText(file, format);
        }

        /// <summary>
        /// Exports the table to a <see cref="Stream"/>
        /// </summary>
        /// <param name="stream">Data will be exported to this stream</param>
        /// <param name="format">Export options</param>
        /// <seealso cref="ExcelRangeBase.SaveToText(Stream, ExcelOutputTextFormat)"></seealso>
        public void SaveToText(Stream stream, ExcelOutputTextFormat format)
        {
            this.Range.SaveToText(stream, format);
        }
#if !NET35 && !NET40
        /// <summary>
        /// Exports the table to a <see cref="Stream"/>
        /// </summary>
        /// <param name="stream">Data will be exported to this stream</param>
        /// <param name="format">Export options</param>
        /// <seealso cref="ExcelRangeBase.SaveToText(Stream, ExcelOutputTextFormat)"></seealso>
        public async Task SaveToTextAsync(Stream stream, ExcelOutputTextFormat format)
        {
            await this.Range.SaveToTextAsync(stream, format);
        }

        /// <summary>
        /// Exports the table to a file
        /// </summary>
        /// <param name="file">Data will be exported to this stream</param>
        /// <param name="format">Export options</param>
        /// <seealso cref="ExcelRangeBase.SaveToTextAsync(FileInfo, ExcelOutputTextFormat)"/>
        public async Task SaveToTextAsync(FileInfo file, ExcelOutputTextFormat format)
        {
            await this.Range.SaveToTextAsync(file, format);
        }

        /// <summary>
        /// Save the table to json
        /// </summary>
        /// <param name="stream">The stream to save to.</param>
        /// <returns></returns>
        public async Task SaveToJsonAsync(Stream stream)
        {
            JsonTableExportSettings? s = new JsonTableExportSettings();
            await this.SaveToJsonInternalAsync(stream, s);
        }

        /// <summary>
        /// Save the table to json
        /// </summary>
        /// <param name="stream">The stream to save to.</param>
        /// <param name="settings">Settings for the json output.</param>
        /// <returns></returns>
        public async Task SaveToJsonAsync(Stream stream, Action<JsonTableExportSettings> settings)
        {
            JsonTableExportSettings? s = new JsonTableExportSettings();
            settings.Invoke(s);
            await this.SaveToJsonInternalAsync(stream, s);
        }

        private async Task SaveToJsonInternalAsync(Stream stream, JsonTableExportSettings s)
        {
            JsonTableExport? exporter = new JsonTableExport(this, s);
            await exporter.ExportAsync(stream);
            await stream.FlushAsync();
        }
#endif

        /// <summary>
        /// Exports the table to a <see cref="System.Data.DataTable"/>
        /// </summary>
        /// <returns>A <see cref="System.Data.DataTable"/> containing the data in the table range</returns>
        /// <seealso cref="ExcelRangeBase.ToDataTable()"/>
        public DataTable ToDataTable()
        {
            return this.Range.ToDataTable();
        }

        /// <summary>
        /// Returns the table as a JSON string
        /// </summary>
        /// <returns>A string containing the JSON document.</returns>
        public string ToJson()
        {
            JsonTableExportSettings? s = new JsonTableExportSettings();

            return this.ToJsonString(s);
        }

        /// <summary>
        /// Returns the table as a JSON string
        /// </summary>
        /// <param name="settings">Settings to configure the JSON output</param>
        /// <returns>A string containing the JSON document.</returns>
        public string ToJson(Action<JsonTableExportSettings> settings)
        {
            JsonTableExportSettings? s = new JsonTableExportSettings();
            settings.Invoke(s);

            return this.ToJsonString(s);
        }

        /// <summary>
        /// Saves the table as a JSON string to a string
        /// </summary>
        /// <param name="stream">The stream to write the JSON to.</param>
        public void SaveToJson(Stream stream)
        {
            JsonTableExportSettings? s = new JsonTableExportSettings();
            this.SaveToJsonInternal(stream, s);
        }

        /// <summary>
        /// Saves the table as a JSON string to a string
        /// </summary>
        /// <param name="stream">The stream to write the JSON to.</param>
        /// <param name="settings">Settings to configure the JSON output</param>
        public void SaveToJson(Stream stream, Action<JsonTableExportSettings> settings)
        {
            JsonTableExportSettings? s = new JsonTableExportSettings();
            settings.Invoke(s);
            this.SaveToJsonInternal(stream, s);
        }

        private void SaveToJsonInternal(Stream stream, JsonTableExportSettings s)
        {
            JsonTableExport? exporter = new JsonTableExport(this, s);
            exporter.Export(stream);
            stream.Flush();
        }

        private string ToJsonString(JsonTableExportSettings s)
        {
            JsonTableExport? exporter = new JsonTableExport(this, s);
            MemoryStream? ms = RecyclableMemory.GetStream();
            exporter.Export(ms);

            return s.Encoding.GetString(ms.ToArray());
        }

        /// <summary>
        /// Exports the table to a <see cref="System.Data.DataTable"/>
        /// </summary>
        /// <returns>A <see cref="System.Data.DataTable"/> containing the data in the table range</returns>
        /// <seealso cref="ExcelRangeBase.ToDataTable(ToDataTableOptions)"/>
        public DataTable ToDataTable(ToDataTableOptions options)
        {
            return this.Range.ToDataTable(options);
        }

        /// <summary>
        /// Exports the table to a <see cref="System.Data.DataTable"/>
        /// </summary>
        /// <returns>A <see cref="System.Data.DataTable"/> containing the data in the table range</returns>
        /// <seealso cref="ExcelRangeBase.ToDataTable(Action{ToDataTableOptions})"/>
        public DataTable ToDataTable(Action<ToDataTableOptions> configHandler)
        {
            return this.Range.ToDataTable(configHandler);
        }

        #endregion

#if (!NET35)
        /// <summary>
        /// Returns a collection of T for the tables data range. The total row is not included.
        /// The table must have headers.
        /// Headers will be mapped to properties using the name or the objects attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <returns>A list of T</returns>
        public List<T> ToCollection<T>()
        {
            if (this.ShowHeader == false)
            {
                throw new InvalidOperationException("The table must have headers.");
            }

            return this.ToCollection<T>(ToCollectionTableOptions.Default);
        }

        /// <summary>
        /// Returns a collection of T for the tables data range. The total row is not included.
        /// The table must have headers.
        /// Headers will be mapped to properties using the name or the property attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <param name="options">Configures the settings for the function</param>
        /// <returns>A list of T</returns>
        public List<T> ToCollection<T>(Action<ToCollectionTableOptions> options)
        {
            ToCollectionTableOptions? o = new ToCollectionTableOptions();
            options.Invoke(o);

            return this.ToCollection<T>(o);
        }

        /// <summary>
        /// Returns a collection of T for the tables data range. The total row is not included.
        /// The table must have headers.
        /// Headers will be mapped to properties using the name or the property attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <param name="options">Settings for the method</param>
        /// <returns>A list of T</returns>
        public List<T> ToCollection<T>(ToCollectionTableOptions options)
        {
            if (this.ShowHeader == false && (options.Headers == null || options.Headers.Length == 0))
            {
                throw new InvalidOperationException("The table must have headers or the headers must be supplied in the options.");
            }

            ToCollectionRangeOptions? ro = new ToCollectionRangeOptions(options) { HeaderRow = 0 };

            if (this.ShowTotal)
            {
                ExcelRangeBase? r = this.Range;

                return this.WorkSheet.Cells[r._fromRow, r._fromCol, r._toRow - 1, r._toCol].ToCollection<T>(ro);
            }
            else
            {
                return this.Range.ToCollection<T>(ro);
            }
        }

#endif
        /// <summary>
        /// Returns a collection of T for the table. 
        /// If the range contains multiple addresses the first range is used.
        /// The the table must have headers.
        /// Headers will be mapped to properties using the name or the attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <param name="setRow">The call back function to map each row to the item of type T.</param>
        /// <returns>A list of T</returns>
        public List<T> ToCollection<T>(Func<Export.ToCollection.ToCollectionRow, T> setRow)
        {
            return this.ToCollectionWithMappings(setRow, ToCollectionTableOptions.Default);
        }

        /// <summary>
        /// Returns a collection of T for the table. 
        /// If the range contains multiple addresses the first range is used.
        /// The the table must have headers.
        /// Headers will be mapped to properties using the name or the attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <param name="setRow">The call back function to map each row to the item of type T.</param>
        /// <param name="options">Configures the settings for the function</param>
        /// <returns>A list of T</returns>
        public List<T> ToCollectionWithMappings<T>(Func<Export.ToCollection.ToCollectionRow, T> setRow, Action<ToCollectionTableOptions> options)
        {
            ToCollectionTableOptions? o = ToCollectionTableOptions.Default;
            options.Invoke(o);

            return this.ToCollectionWithMappings(setRow, o);
        }

        /// <summary>
        /// Returns a collection of T for the table. 
        /// If the range contains multiple addresses the first range is used.
        /// The the table must have headers.
        /// Headers will be mapped to properties using the name or the attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <param name="setRow">The call back function to map each row to the item of type T.</param>
        /// <param name="options">Configures the settings for the function</param>
        /// <returns>A list of T</returns>
        public List<T> ToCollectionWithMappings<T>(Func<Export.ToCollection.ToCollectionRow, T> setRow, ToCollectionTableOptions options)
        {
            if (this.ShowHeader == false && (options.Headers == null || options.Headers.Length == 0))
            {
                throw new InvalidOperationException("The table must have headers or the headers must be supplied in the options.");
            }

            ToCollectionRangeOptions? ro = new ToCollectionRangeOptions(options);
            ro.HeaderRow = 0;

            if (this.ShowTotal)
            {
                ExcelRangeBase? r = this.Range;

                return this.WorkSheet.Cells[r._fromRow, r._fromCol, r._toRow - 1, r._toCol].ToCollectionWithMappings(setRow, ro);
            }
            else
            {
                return this.Range.ToCollectionWithMappings(setRow, ro);
            }
        }

        internal ExcelTableColumnCollection _cols;

        /// <summary>
        /// Collection of the columns in the table
        /// </summary>
        public ExcelTableColumnCollection Columns
        {
            get { return this._cols ??= new ExcelTableColumnCollection(this); }
        }

        TableStyles _tableStyle = TableStyles.Medium6;

        /// <summary>
        /// The table style. If this property is custom, the style from the StyleName propery is used.
        /// </summary>
        public TableStyles TableStyle
        {
            get { return this._tableStyle; }
            set
            {
                this._tableStyle = value;

                if (value != TableStyles.Custom)
                {
                    this.SetXmlNodeString(STYLENAME_PATH, "TableStyle" + value.ToString());
                }
            }
        }

        const string HEADERROWCOUNT_PATH = "@headerRowCount";
        const string AUTOFILTER_PATH = "d:autoFilter";
        const string AUTOFILTER_ADDRESS_PATH = AUTOFILTER_PATH + "/@ref";

        /// <summary>
        /// If the header row is visible or not
        /// </summary>
        public bool ShowHeader
        {
            get { return this.GetXmlNodeInt(HEADERROWCOUNT_PATH) != 0; }
            set
            {
                if ((this.Address._toRow - this.Address._fromRow < 0 && value) || (this.Address._toRow - this.Address._fromRow == 1 && value && this.ShowTotal))
                {
                    throw new Exception("Cant set ShowHeader-property. Table has too few rows");
                }

                if (value)
                {
                    this.DeleteNode(HEADERROWCOUNT_PATH);
                    this.WriteAutoFilter(this.ShowTotal);

                    for (int i = 0; i < this.Columns.Count; i++)
                    {
                        string? v = this.WorkSheet.GetValue<string>(this.Address._fromRow, this.Address._fromCol + i);

                        if (string.IsNullOrEmpty(v))
                        {
                            this.WorkSheet.SetValue(this.Address._fromRow, this.Address._fromCol + i, this._cols[i].Name);
                        }
                        else if (v != this._cols[i].Name)
                        {
                            this._cols[i].Name = v;
                        }
                    }

                    this.HeaderRowStyle.SetStyle();

                    foreach (ExcelTableColumn? c in this.Columns)
                    {
                        c.HeaderRowStyle.SetStyle();
                    }
                }
                else
                {
                    this.SetXmlNodeString(HEADERROWCOUNT_PATH, "0");
                    this.DeleteAllNode(AUTOFILTER_ADDRESS_PATH);
                    this.DataStyle.SetStyle();
                }
            }
        }

        internal ExcelAddressBase AutoFilterAddress
        {
            get
            {
                string a = this.GetXmlNodeString(AUTOFILTER_ADDRESS_PATH);

                if (a == "")
                {
                    return null;
                }
                else
                {
                    return new ExcelAddressBase(a);
                }
            }
        }

        ExcelAutoFilter _autoFilter;

        /// <summary>
        /// Autofilter settings for the table
        /// </summary>
        public ExcelAutoFilter AutoFilter
        {
            get
            {
                if (this.ShowFilter)
                {
                    return this._autoFilter;
                }
                else
                {
                    return null;
                }
            }
        }

        private void WriteAutoFilter(bool showTotal)
        {
            if (this.ShowHeader)
            {
                string autofilterAddress;

                if (showTotal)
                {
                    autofilterAddress = ExcelCellBase.GetAddress(this.Address._fromRow, this.Address._fromCol, this.Address._toRow - 1, this.Address._toCol);
                }
                else
                {
                    autofilterAddress = this.Address.Address;
                }

                this.SetXmlNodeString(AUTOFILTER_ADDRESS_PATH, autofilterAddress);
                this.SetAutoFilter();
            }
        }

        private void SetAutoFilter()
        {
            if (this._autoFilter == null)
            {
                XmlNode? node = this.TopNode.SelectSingleNode(AUTOFILTER_PATH, this.NameSpaceManager);
                this._autoFilter = new ExcelAutoFilter(this.NameSpaceManager, node, this);
                this._autoFilter.Address = this.AutoFilterAddress;
            }
        }

        /// <summary>
        /// If the header row has an autofilter
        /// </summary>
        public bool ShowFilter
        {
            get { return this.ShowHeader && this.AutoFilterAddress != null; }
            set
            {
                if (this.ShowHeader)
                {
                    if (value)
                    {
                        this.WriteAutoFilter(this.ShowTotal);
                    }
                    else
                    {
                        this.DeleteAllNode(AUTOFILTER_PATH);
                        this._autoFilter = null;
                    }
                }
                else if (value)
                {
                    throw new InvalidOperationException("Filter can only be applied when ShowHeader is set to true");
                }
            }
        }

        const string TOTALSROWCOUNT_PATH = "@totalsRowCount";
        const string TOTALSROWSHOWN_PATH = "@totalsRowShown";

        /// <summary>
        /// If the total row is visible or not
        /// </summary>
        public bool ShowTotal
        {
            get { return this.GetXmlNodeInt(TOTALSROWCOUNT_PATH) == 1; }
            set
            {
                if (value != this.ShowTotal)
                {
                    if (value)
                    {
                        this.Address = new ExcelAddress(this.WorkSheet.Name,
                                                        ExcelCellBase.GetAddress(this.Address.Start.Row,
                                                                                 this.Address.Start.Column,
                                                                                 this.Address.End.Row + 1,
                                                                                 this.Address.End.Column));
                    }
                    else
                    {
                        this.Address = new ExcelAddress(this.WorkSheet.Name,
                                                        ExcelCellBase.GetAddress(this.Address.Start.Row,
                                                                                 this.Address.Start.Column,
                                                                                 this.Address.End.Row - 1,
                                                                                 this.Address.End.Column));
                    }

                    this.SetXmlNodeString("@ref", this.Address.Address);

                    if (value)
                    {
                        this.SetXmlNodeString(TOTALSROWCOUNT_PATH, "1");
                        this.SetXmlNodeString(TOTALSROWSHOWN_PATH, "1");
                        this.TotalsRowStyle.SetStyle();

                        foreach (ExcelTableColumn? c in this.Columns)
                        {
                            c.TotalsRowStyle.SetStyle();
                        }
                    }
                    else
                    {
                        this.DeleteNode(TOTALSROWCOUNT_PATH);
                        this.DataStyle.SetStyle();
                    }

                    this.WriteAutoFilter(value);
                }
            }
        }

        const string STYLENAME_PATH = "d:tableStyleInfo/@name";

        /// <summary>
        /// The style name for custum styles
        /// </summary>
        public string StyleName
        {
            get { return this.GetXmlNodeString(STYLENAME_PATH); }
            set
            {
                this._tableStyle = GetTableStyle(value);

                if (this._tableStyle == TableStyles.None)
                {
                    this.DeleteAllNode(STYLENAME_PATH);
                }
                else
                {
                    this.SetXmlNodeString(STYLENAME_PATH, value);
                }
            }
        }

        private static TableStyles GetTableStyle(string value)
        {
            if (value.StartsWith("TableStyle"))
            {
                try
                {
                    return (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10, value.Length - 10), true);
                }
                catch
                {
                    return TableStyles.Custom;
                }
            }
            else if (value == "None")
            {
                return TableStyles.None;
            }
            else
            {
                return TableStyles.Custom;
            }
        }

        const string SHOWFIRSTCOLUMN_PATH = "d:tableStyleInfo/@showFirstColumn";

        /// <summary>
        /// Display special formatting for the first row
        /// </summary>
        public bool ShowFirstColumn
        {
            get { return this.GetXmlNodeBool(SHOWFIRSTCOLUMN_PATH); }
            set { this.SetXmlNodeBool(SHOWFIRSTCOLUMN_PATH, value, false); }
        }

        const string SHOWLASTCOLUMN_PATH = "d:tableStyleInfo/@showLastColumn";

        /// <summary>
        /// Display special formatting for the last row
        /// </summary>
        public bool ShowLastColumn
        {
            get { return this.GetXmlNodeBool(SHOWLASTCOLUMN_PATH); }
            set { this.SetXmlNodeBool(SHOWLASTCOLUMN_PATH, value, false); }
        }

        const string SHOWROWSTRIPES_PATH = "d:tableStyleInfo/@showRowStripes";

        /// <summary>
        /// Display banded rows
        /// </summary>
        public bool ShowRowStripes
        {
            get { return this.GetXmlNodeBool(SHOWROWSTRIPES_PATH); }
            set { this.SetXmlNodeBool(SHOWROWSTRIPES_PATH, value, false); }
        }

        const string SHOWCOLUMNSTRIPES_PATH = "d:tableStyleInfo/@showColumnStripes";

        /// <summary>
        /// Display banded columns
        /// </summary>
        public bool ShowColumnStripes
        {
            get { return this.GetXmlNodeBool(SHOWCOLUMNSTRIPES_PATH); }
            set { this.SetXmlNodeBool(SHOWCOLUMNSTRIPES_PATH, value, false); }
        }

        const string TOTALSROWCELLSTYLE_PATH = "@totalsRowCellStyle";

        /// <summary>
        /// Named style used for the total row
        /// </summary>
        public string TotalsRowCellStyle
        {
            get { return this.GetXmlNodeString(TOTALSROWCELLSTYLE_PATH); }
            set
            {
                if (this.WorkSheet.Workbook.Styles.NamedStyles.FindIndexById(value) < 0)
                {
                    throw new Exception(string.Format("Named style {0} does not exist.", value));
                }

                this.SetXmlNodeString(this.TopNode, TOTALSROWCELLSTYLE_PATH, value, true);

                if (this.ShowTotal)
                {
                    this.WorkSheet.Cells[this.Address._toRow, this.Address._fromCol, this.Address._toRow, this.Address._toCol].StyleName = value;
                }
            }
        }

        const string DATACELLSTYLE_PATH = "@dataCellStyle";

        /// <summary>
        /// Named style used for the data cells
        /// </summary>
        public string DataCellStyleName
        {
            get { return this.GetXmlNodeString(DATACELLSTYLE_PATH); }
            set
            {
                if (this.WorkSheet.Workbook.Styles.NamedStyles.FindIndexById(value) < 0)
                {
                    throw new Exception(string.Format("Named style {0} does not exist.", value));
                }

                this.SetXmlNodeString(this.TopNode, DATACELLSTYLE_PATH, value, true);

                int fromRow = this.Address._fromRow + (this.ShowHeader ? 1 : 0),
                    toRow = this.Address._toRow - (this.ShowTotal ? 1 : 0);

                if (fromRow < toRow)
                {
                    this.WorkSheet.Cells[fromRow, this.Address._fromCol, toRow, this.Address._toCol].StyleName = value;
                }
            }
        }

        const string HEADERROWCELLSTYLE_PATH = "@headerRowCellStyle";

        /// <summary>
        /// Named style used for the header row
        /// </summary>
        public string HeaderRowCellStyle
        {
            get { return this.GetXmlNodeString(HEADERROWCELLSTYLE_PATH); }
            set
            {
                if (this.WorkSheet.Workbook.Styles.NamedStyles.FindIndexById(value) < 0)
                {
                    throw new Exception(string.Format("Named style {0} does not exist.", value));
                }

                this.SetXmlNodeString(this.TopNode, HEADERROWCELLSTYLE_PATH, value, true);

                if (this.ShowHeader)
                {
                    this.WorkSheet.Cells[this.Address._fromRow, this.Address._fromCol, this.Address._fromRow, this.Address._toCol].StyleName = value;
                }
            }
        }

        /// <summary>
        /// Checkes if two tables are the same
        /// </summary>
        /// <param name="x">Table 1</param>
        /// <param name="y">Table 2</param>
        /// <returns></returns>
        public bool Equals(ExcelTable x, ExcelTable y)
        {
            return x.WorkSheet == y.WorkSheet && x.Id == y.Id && x.TableXml.OuterXml == y.TableXml.OuterXml;
        }

        /// <summary>
        /// Returns a hashcode generated from the TableXml
        /// </summary>
        /// <param name="obj">The table</param>
        /// <returns>The hashcode</returns>
        public int GetHashCode(ExcelTable obj)
        {
            return obj.TableXml.OuterXml.GetHashCode();
        }

        /// <summary>
        /// Adds new rows to the table. 
        /// </summary>
        /// <param name="rows">Number of rows to add to the table. Default is 1</param>
        /// <returns></returns>
        public ExcelRangeBase AddRow(int rows = 1)
        {
            return this.InsertRow(int.MaxValue, rows);
        }

        /// <summary>
        /// Inserts one or more rows before the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the row will be inserted. Default is in the end of the table. 0 will insert the row at the top. Any value larger than the number of rows in the table will insert a row at the bottom of the table.</param>
        /// <param name="rows">Number of rows to insert.</param>
        /// <param name="copyStyles">Copy styles from the row above. If inserting a row at position 0, the first row will be used as a template.</param>
        /// <returns>The inserted range</returns>
        public ExcelRangeBase InsertRow(int position, int rows = 1, bool copyStyles = true)
        {
            if (position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }

            if (rows < 0)
            {
                throw new ArgumentException("position", "rows can't be negative");
            }

            int firstRow = this._address._fromRow;
            bool isFirstRow = position == 0;
            int subtact = this.ShowTotal ? 2 : 1;

            if (position >= ExcelPackage.MaxRows || position > this._address._fromRow + position + rows - subtact)
            {
                position = this._address.Rows - subtact;
            }

            if (this._address._fromRow + position + rows > ExcelPackage.MaxRows)
            {
                throw new InvalidOperationException("Insert will exceed the maximum number of rows in the worksheet");
            }

            if (this.ShowHeader)
            {
                position++;
            }

            string? address = ExcelCellBase.GetAddress(this._address._fromRow + position,
                                                       this._address._fromCol,
                                                       this._address._fromRow + position + rows - 1,
                                                       this._address._toCol);

            ExcelRangeBase? range = new ExcelRangeBase(this.WorkSheet, address);

            WorksheetRangeInsertHelper.Insert(range, eShiftTypeInsert.Down, false, true);

            this.ExtendCalculatedFormulas(range);

            if (copyStyles)
            {
                int copyFromRow = isFirstRow ? this.DataRange._fromRow + rows + 1 : this._address._fromRow + position - 1;

                if (range._toRow > this._address._toRow)
                {
                    this.Address = this._address.AddRow(this._address._toRow, rows);
                }

                this.CopyStylesFromRow(address,
                                       copyFromRow); //Separate copy instead of using Insert paramter 3 as the first row should not copy the styles from the header row.
            }

            if (this._address._fromRow > firstRow)
            {
                this._address = new ExcelAddressBase(firstRow,
                                                     this._address._fromCol,
                                                     this._address._toRow,
                                                     this._address._toCol,
                                                     this._address._fromRowFixed,
                                                     this._address._fromColFixed,
                                                     this._address._toRowFixed,
                                                     this._address._toColFixed,
                                                     this._address.WorkSheetName,
                                                     null);
            }

            return range;
        }

        private void ExtendCalculatedFormulas(ExcelRangeBase range)
        {
            foreach (ExcelTableColumn? c in this.Columns)
            {
                if (!string.IsNullOrEmpty(c.CalculatedColumnFormula))
                {
                    c.SetFormulaCells(range._fromRow, range._toRow, range._fromCol + c.Position);
                }
            }
        }

        private void CopyStylesFromRow(string address, int copyRow)
        {
            ExcelRange? range = this.WorkSheet.Cells[address];

            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                int styleId = this.WorkSheet.Cells[copyRow, col].StyleID;

                if (styleId != 0)
                {
                    for (int row = range._fromRow; row <= range._toRow; row++)
                    {
                        this.WorkSheet.SetStyleInner(row, col, styleId);
                    }
                }
            }
        }

        private void CopyStylesFromColumn(string address, int copyColumn)
        {
            ExcelRange? range = this.WorkSheet.Cells[address];

            for (int row = range._fromRow; row <= range._toRow; row++)
            {
                int styleId = this.WorkSheet.Cells[row, copyColumn].StyleID;

                if (styleId != 0)
                {
                    for (int col = range._fromCol; col <= range._toCol; col++)
                    {
                        this.WorkSheet.SetStyleInner(row, col, styleId);
                    }
                }
            }
        }

        /// <summary>
        /// Deletes one or more rows at the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the row will be deleted. 0 will delete the first row. </param>
        /// <param name="rows">Number of rows to delete.</param>
        /// <returns></returns>
        public ExcelRangeBase DeleteRow(int position, int rows = 1)
        {
            if (position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }

            if (rows < 0)
            {
                throw new ArgumentException("position", "rows can't be negative");
            }

            if (this._address._fromRow + position + rows + (this.ShowHeader ? 0 : -1) > this._address._toRow)
            {
                throw new InvalidOperationException("Delete will exceed the number of rows in the table");
            }

            int subtract = this.ShowTotal ? 2 : 1;

            if (position == 0 && rows + subtract >= this._address.Rows)
            {
                throw new InvalidOperationException("Can't delete all table rows. A table must have at least one row.");
            }

            position += this.ShowHeader ? 1 : 0; //Header row should not be deleted.

            string? address = ExcelCellBase.GetAddress(this._address._fromRow + position,
                                                       this._address._fromCol,
                                                       this._address._fromRow + position + rows - 1,
                                                       this._address._toCol);

            ExcelRangeBase? range = new ExcelRangeBase(this.WorkSheet, address);
            range.Delete(eShiftTypeDelete.Up);

            return range;
        }

        /// <summary>
        /// Inserts one or more columns before the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the column will be inserted. 0 will insert the column at the leftmost. Any value larger than the number of rows in the table will insert a row at the bottom of the table.</param>
        /// <param name="columns">Number of rows to insert.</param>
        /// <param name="copyStyles">Copy styles from the column to the left.</param>
        /// <returns>The inserted range</returns>
        internal ExcelRangeBase InsertColumn(int position, int columns, bool copyStyles = false)
        {
            if (position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }

            if (columns < 0)
            {
                throw new ArgumentException("columns", "columns can't be negative");
            }

            bool isFirstColumn = position == 0;

            if (position >= ExcelPackage.MaxColumns || position > this._address._fromCol + position + columns - 1)
            {
                position = this._address.Columns;
            }

            if (this._address._fromCol + position + columns - 1 > ExcelPackage.MaxColumns)
            {
                throw new InvalidOperationException("Insert will exceed the maximum number of columns in the worksheet");
            }

            string? address = ExcelCellBase.GetAddress(this._address._fromRow,
                                                       this._address._fromCol + position,
                                                       this._address._toRow,
                                                       this._address._fromCol + position + columns - 1);

            ExcelRangeBase? range = new ExcelRangeBase(this.WorkSheet, address);

            WorksheetRangeInsertHelper.Insert(range, eShiftTypeInsert.Right, true, false);

            if (position == 0)
            {
                this.Address = new ExcelAddressBase(this._address._fromRow, this._address._fromCol - columns, this._address._toRow, this._address._toCol);
            }
            else if (range._toCol > this._address._toCol)
            {
                this.Address = new ExcelAddressBase(this._address._fromRow, this._address._fromCol, this._address._toRow, this._address._toCol + columns);
            }

            if (copyStyles && isFirstColumn == false)
            {
                int copyFromCol = this._address._fromCol + position - 1;
                this.CopyStylesFromColumn(address, copyFromCol);
            }

            return range;
        }

        /// <summary>
        /// Deletes one or more columns at the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the column will be deleted.</param>
        /// <param name="columns">Number of rows to delete.</param>
        /// <returns>The deleted range</returns>
        internal ExcelRangeBase DeleteColumn(int position, int columns)
        {
            if (position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }

            if (columns < 0)
            {
                throw new ArgumentException("columns", "columns can't be negative");
            }

            if (this._address._toCol < this._address._fromCol + position + columns - 1)
            {
                throw new InvalidOperationException("Delete will exceed the number of columns in the table");
            }

            string? address = ExcelCellBase.GetAddress(this._address._fromRow,
                                                       this._address._fromCol + position,
                                                       this._address._toRow,
                                                       this._address._fromCol + position + columns - 1);

            ExcelRangeBase? range = new ExcelRangeBase(this.WorkSheet, address);

            WorksheetRangeDeleteHelper.Delete(range, eShiftTypeDelete.Left);

            return range;
        }

        internal int? HeaderRowBorderDxfId
        {
            get { return this.GetXmlNodeIntNull("@headerRowBorderDxfId"); }
            set { this.SetXmlNodeInt("@headerRowBorderDxfId", value); }
        }

        /// <summary>
        /// Sets differential formatting styles for the table header row border style.
        /// </summary>
        public ExcelDxfBorderBase HeaderRowBorderStyle { get; set; }

        internal int? TableBorderDxfId
        {
            get { return this.GetXmlNodeIntNull("@tableBorderDxfId"); }
            set { this.SetXmlNodeInt("@tableBorderDxfId", value); }
        }

        /// <summary>
        /// Sets differential formatting styles for the tables row border style.
        /// </summary>
        public ExcelDxfBorderBase TableBorderStyle { get; set; }

        #region Sorting

        private TableSorter _tableSorter;
        const string SortStatePath = "d:sortState";
        SortState _sortState;

        /// <summary>
        /// Gets the sort state of the table.
        /// <seealso cref="Sort(Action{TableSortOptions})"/>
        /// <seealso cref="Sort(TableSortOptions)"/>
        /// </summary>
        public SortState SortState
        {
            get
            {
                if (this._sortState == null)
                {
                    XmlNode? node = this.TableXml.SelectSingleNode($"//{SortStatePath}", this.NameSpaceManager);

                    if (node == null)
                    {
                        return null;
                    }

                    this._sortState = new SortState(this.NameSpaceManager, node);
                }

                return this._sortState;
            }
        }

        internal void SetTableSortState(int[] columns, bool[] descending, CompareOptions compareOptions, Dictionary<int, string[]> customLists)
        {
            //Set sort state
            SortState? sortState = new SortState(this.Range.Worksheet.NameSpaceManager, this);
            sortState.Clear();
            ExcelRangeBase? dataRange = this.DataRange;
            sortState.Ref = dataRange.Address;
            sortState.CaseSensitive = compareOptions == CompareOptions.IgnoreCase || compareOptions == CompareOptions.OrdinalIgnoreCase;

            for (int ix = 0; ix < columns.Length; ix++)
            {
                bool? desc = null;

                if (descending.Length > ix && descending[ix])
                {
                    desc = true;
                }

                string? adr =
                    ExcelCellBase.GetAddress(dataRange._fromRow, dataRange._fromCol + columns[ix], dataRange._toRow, dataRange._fromCol + columns[ix]);

                if (customLists.ContainsKey(columns[ix]))
                {
                    sortState.SortConditions.Add(adr, desc, customLists[columns[ix]]);
                }
                else
                {
                    sortState.SortConditions.Add(adr, desc);
                }
            }
        }

        /// <summary>
        /// Sorts the data in the table according to the supplied <see cref="RangeSortOptions"/>
        /// </summary>
        /// <param name="options"></param>
        /// <example> 
        /// <code>
        /// var options = new SortOptions();
        /// options.SortBy.Column(0).ThenSortBy.Column(1, eSortDirection.Descending);
        /// </code>
        /// </example>
        public void Sort(TableSortOptions options)
        {
            this._tableSorter.Sort(options);
        }

        /// <summary>
        /// Sorts the data in the table according to the supplied action of <see cref="RangeSortOptions"/>
        /// </summary>
        /// <example> 
        /// <code>
        /// table.Sort(x =&gt; x.SortBy.Column(0).ThenSortBy.Column(1, eSortDirection.Descending);
        /// </code>
        /// </example>
        /// <param name="configuration">An action with parameters for sorting</param>
        public void Sort(Action<TableSortOptions> configuration)
        {
            this._tableSorter.Sort(configuration);
        }

        #endregion
    }
}