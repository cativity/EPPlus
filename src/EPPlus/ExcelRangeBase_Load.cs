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

using OfficeOpenXml.Attributes;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.LoadFunctions;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml
{
    public partial class ExcelRangeBase
    {
        #region LoadFromDataReader

        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to loadfrom</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableName">The name of the table</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders, string TableName, TableStyles TableStyle = TableStyles.None)
        {
            ExcelRangeBase? r = this.LoadFromDataReader(Reader, PrintHeaders);

            int rows = r.Rows - 1;

            if (rows >= 0 && r.Columns > 0)
            {
                ExcelTable? tbl =
                    this._worksheet.Tables.Add(new ExcelAddressBase(this._fromRow,
                                                                    this._fromCol,
                                                                    this._fromRow + (rows <= 0 ? 1 : rows),
                                                                    this._fromCol + r.Columns - 1),
                                               TableName);

                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }

            return r;
        }

        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders)
        {
            if (Reader == null)
            {
                throw new ArgumentNullException(nameof(Reader), "Reader can't be null");
            }

            int fieldCount = Reader.FieldCount;

            int col = this._fromCol,
                row = this._fromRow;

            if (PrintHeaders)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    // If no caption is set, the ColumnName property is called implicitly.
                    this._worksheet.SetValueInner(row, col++, Reader.GetName(i));
                }

                row++;
                col = this._fromCol;
            }

            while (Reader.Read())
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    this._worksheet.SetValueInner(row, col++, Reader.GetValue(i));
                }

                row++;
                col = this._fromCol;
            }

            return this._worksheet.Cells[this._fromRow, this._fromCol, row - 1, this._fromCol + fieldCount - 1];
        }
#if !NET35 && !NET40
        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to loadfrom</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableName">The name of the table</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <param name="cancellationToken">The cancellation token to use</param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader,
                                                                  bool PrintHeaders,
                                                                  string TableName,
                                                                  TableStyles TableStyle = TableStyles.None,
                                                                  CancellationToken? cancellationToken = null)
        {
            cancellationToken ??= CancellationToken.None;
            ExcelRangeBase? r = await this.LoadFromDataReaderAsync(Reader, PrintHeaders, cancellationToken.Value).ConfigureAwait(false);

            if (cancellationToken.Value.IsCancellationRequested)
            {
                return r;
            }

            int rows = r.Rows - 1;

            if (rows >= 0 && r.Columns > 0)
            {
                ExcelTable? tbl =
                    this._worksheet.Tables.Add(new ExcelAddressBase(this._fromRow,
                                                                    this._fromCol,
                                                                    this._fromRow + (rows <= 0 ? 1 : rows),
                                                                    this._fromCol + r.Columns - 1),
                                               TableName);

                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }

            return r;
        }

        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader, bool PrintHeaders) => await this.LoadFromDataReaderAsync(Reader, PrintHeaders, CancellationToken.None);

        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="cancellationToken">The cancellation token to use</param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader, bool PrintHeaders, CancellationToken cancellationToken)
        {
            if (Reader == null)
            {
                throw new ArgumentNullException(nameof(Reader), "Reader can't be null");
            }

            int fieldCount = Reader.FieldCount;

            int col = this._fromCol,
                row = this._fromRow;

            if (PrintHeaders)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    // If no caption is set, the ColumnName property is called implicitly.
                    this._worksheet.SetValueInner(row, col++, Reader.GetName(i));
                }

                row++;
                col = this._fromCol;
            }

            while (await Reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    this._worksheet.SetValueInner(row, col++, Reader.GetValue(i));
                }

                row++;
                col = this._fromCol;

                if (row % 100 == 0 && cancellationToken.IsCancellationRequested) //Check every 100 rows
                {
                    return this._worksheet.Cells[this._fromRow, this._fromCol, row - 1, this._fromCol + fieldCount - 1];
                }
            }

            return this._worksheet.Cells[this._fromRow, this._fromCol, row - 1, this._fromCol + fieldCount - 1];
        }
#endif

        #endregion

        #region LoadFromDataTable

        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders, TableStyles? TableStyle)
        {
            LoadFromDataTableParams? parameters = new LoadFromDataTableParams { PrintHeaders = PrintHeaders, TableStyle = TableStyle };
            LoadFromDataTable? func = new LoadFromDataTable(this, Table, parameters);

            return func.Load();
        }

        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders) => this.LoadFromDataTable(Table, PrintHeaders, null);

        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="table">The datatable to load</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable table) => this.LoadFromDataTable(table, false, null);

        /// <summary>
        /// Load the data from the <see cref="DataTable"/> starting from the top left cell of the range
        /// </summary>
        /// <param name="table"></param>
        /// <param name="paramsConfig"><see cref="Action{LoacFromCollectionParams}"/> to provide parameters to the function</param>
        /// <example>
        /// <code>
        /// sheet.Cells["C1"].LoadFromDataTable(dataTable, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </code>
        /// </example>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable table, Action<LoadFromDataTableParams> paramsConfig)
        {
            LoadFromDataTableParams? parameters = new LoadFromDataTableParams();
            paramsConfig.Invoke(parameters);

            return this.LoadFromDataTable(table, parameters.PrintHeaders, parameters.TableStyle);
        }

        #endregion

        #region LoadFromArrays

        /// <summary>
        /// Loads data from the collection of arrays of objects into the range, starting from
        /// the top-left cell.
        /// </summary>
        /// <param name="Data">The data.</param>
        public ExcelRangeBase LoadFromArrays(IEnumerable<object[]> Data)
        {
            //thanx to Abdullin for the code contribution
            if (!(Data?.Any() ?? false))
            {
                return null;
            }

            int maxColumn = 0;
            int row = this._fromRow;

            foreach (object[] item in Data)
            {
                this._worksheet._values.SetValueRow_Value(row, this._fromCol, item);

                if (maxColumn < item.Length)
                {
                    maxColumn = item.Length;
                }

                row++;
            }

            return this._worksheet.Cells[this._fromRow, this._fromCol, row - 1, this._fromCol + maxColumn - 1];
        }

        #endregion

        #region LoadFromCollection

        /// <summary>
        /// Load a collection into a the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection)
        {
            Type? type = typeof(T);
            EpplusTableAttribute? attr = type.GetFirstAttributeOfType<EpplusTableAttribute>();

            if (attr != null)
            {
                ExcelRangeBase? range =
                    this.LoadFromCollection(Collection, attr.PrintHeaders, attr.TableStyle, BindingFlags.Public | BindingFlags.Instance, null);

                if (attr.AutofitColumns)
                {
                    range.AutoFitColumns();
                }

                if (attr.AutoCalculate)
                {
                    range.Calculate();
                }

                return range;
            }

            return this.LoadFromCollection<T>(Collection, false, null, BindingFlags.Public | BindingFlags.Instance, null);
        }

        /// <summary>
        /// Load a collection of T into the worksheet starting from the top left row of the range.
        /// Default option will load all public instance properties of T
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders) => this.LoadFromCollection<T>(Collection, PrintHeaders, null, BindingFlags.Public | BindingFlags.Instance, null);

        /// <summary>
        /// Load a collection of T into the worksheet starting from the top left row of the range.
        /// Default option will load all public instance properties of T
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles? TableStyle) => this.LoadFromCollection<T>(Collection, PrintHeaders, TableStyle, BindingFlags.Public | BindingFlags.Instance, null);

        /// <summary>
        /// Load a collection into the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. Any underscore in the property name will be converted to a space. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="memberFlags">Property flags to use</param>
        /// <param name="Members">The properties to output. Must be of type T</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection,
                                                    bool PrintHeaders,
                                                    TableStyles? TableStyle,
                                                    BindingFlags memberFlags,
                                                    MemberInfo[] Members) =>
            this.LoadFromCollectionInternal(Collection, PrintHeaders, TableStyle, memberFlags, Members);

        private ExcelRangeBase LoadFromCollectionInternal<T>(IEnumerable<T> Collection,
                                                             bool PrintHeaders,
                                                             TableStyles? TableStyle,
                                                             BindingFlags memberFlags,
                                                             MemberInfo[] Members)
        {
            if (Collection is IEnumerable<IDictionary<string, object>>)
            {
                if (Members == null)
                {
                    return this.LoadFromDictionaries(Collection as IEnumerable<IDictionary<string, object>>, PrintHeaders, TableStyle);
                }

                return this.LoadFromDictionaries(Collection as IEnumerable<IDictionary<string, object>>, PrintHeaders, TableStyle, Members.Select(x => x.Name));
            }

            LoadFromCollectionParams? param = new LoadFromCollectionParams
            {
                PrintHeaders = PrintHeaders, TableStyle = TableStyle, BindingFlags = memberFlags, Members = Members
            };

            LoadFromCollection<T>? func = new LoadFromCollection<T>(this, Collection, param);

            return func.Load();
        }

        /// <summary>
        /// Load a collection into the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="collection">The collection to load</param>
        /// <param name="paramsConfig"><see cref="Action{LoacFromCollectionParams}"/> to provide parameters to the function</param>
        /// <example>
        /// <code>
        /// sheet.Cells["C1"].LoadFromCollection(items, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </code>
        /// </example>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> collection, Action<LoadFromCollectionParams> paramsConfig)
        {
            LoadFromCollectionParams? param = new LoadFromCollectionParams();
            paramsConfig.Invoke(param);

            if (collection is IEnumerable<IDictionary<string, object>>)
            {
                if (param.Members == null)
                {
                    return this.LoadFromDictionaries(collection as IEnumerable<IDictionary<string, object>>, param.PrintHeaders, param.TableStyle);
                }

                return this.LoadFromDictionaries(collection as IEnumerable<IDictionary<string, object>>,
                                                 param.PrintHeaders,
                                                 param.TableStyle,
                                                 param.Members.Select(x => x.Name));
            }

            LoadFromCollection<T>? func = new LoadFromCollection<T>(this, collection, param);

            return func.Load();
        }

        #endregion

        #region LoadFromText

        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// Default settings is Comma separation
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <returns>The range containing the data</returns>
        public ExcelRangeBase LoadFromText(string Text) => this.LoadFromText(Text, new ExcelTextFormat());

        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns>The range containing the data</returns>
        public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format)
        {
            if (string.IsNullOrEmpty(Text))
            {
                ExcelRange? r = this._worksheet.Cells[this._fromRow, this._fromCol];
                r.Value = "";

                return r;
            }

            LoadFromTextParams? parameters = new LoadFromTextParams { Format = Format };
            LoadFromText? func = new LoadFromText(this, Text, parameters);

            return func.Load();
        }

        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style. If this parameter is not null no table will be created.</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format, TableStyles? TableStyle, bool FirstRowIsHeader)
        {
            ExcelRangeBase? r = this.LoadFromText(Text, Format);

            if (r != null && TableStyle.HasValue)
            {
                ExcelTable? tbl = this._worksheet.Tables.Add(r, "");
                tbl.ShowHeader = FirstRowIsHeader;
                tbl.TableStyle = TableStyle.Value;
            }

            return r;
        }

        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell using ASCII Encoding.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile) => this.LoadFromText(File.ReadAllText(TextFile.FullName, Encoding.ASCII));

        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format)
        {
            if (TextFile.Exists == false)
            {
                throw new ArgumentException($"File does not exist {TextFile.FullName}");
            }

            return this.LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format);
        }

        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format, TableStyles? TableStyle, bool FirstRowIsHeader)
        {
            if (TextFile.Exists == false)
            {
                throw new ArgumentException($"File does not exist {TextFile.FullName}");
            }

            return this.LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format, TableStyle, FirstRowIsHeader);
        }

        #region LoadFromText async

#if !NET35 && !NET40
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <returns></returns>
        public async Task<ExcelRangeBase> LoadFromTextAsync(FileInfo TextFile)
        {
            if (TextFile.Exists == false)
            {
                throw new ArgumentException($"File does not exist {TextFile.FullName}");
            }

            FileStream? fs = new FileStream(TextFile.FullName, FileMode.Open, FileAccess.Read);
            StreamReader? sr = new StreamReader(fs, Encoding.ASCII);

            return this.LoadFromText(await sr.ReadToEndAsync().ConfigureAwait(false));
        }

        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns></returns>
        public async Task<ExcelRangeBase> LoadFromTextAsync(FileInfo TextFile, ExcelTextFormat Format)
        {
            if (TextFile.Exists == false)
            {
                throw new ArgumentException($"File does not exist {TextFile.FullName}");
            }

            FileStream? fs = new FileStream(TextFile.FullName, FileMode.Open, FileAccess.Read);
            StreamReader? sr = new StreamReader(fs, Format.Encoding);

            return this.LoadFromText(await sr.ReadToEndAsync().ConfigureAwait(false), Format);
        }

        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public async Task<ExcelRangeBase> LoadFromTextAsync(FileInfo TextFile, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
        {
            if (TextFile.Exists == false)
            {
                throw new ArgumentException($"File does not exist {TextFile.FullName}");
            }

            FileStream? fs = new FileStream(TextFile.FullName, FileMode.Open, FileAccess.Read);
            StreamReader? sr = new StreamReader(fs, Format.Encoding);

            return this.LoadFromText(await sr.ReadToEndAsync().ConfigureAwait(false), Format, TableStyle, FirstRowIsHeader);
        }
#endif

        #endregion

        #endregion

        #region LoadFromDictionaries

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items) => this.LoadFromDictionaries(items, false, TableStyles.None, null);

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, bool printHeaders) => this.LoadFromDictionaries(items, printHeaders, TableStyles.None, null);

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, bool printHeaders, TableStyles? tableStyle) => this.LoadFromDictionaries(items, printHeaders, tableStyle, null);

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries</param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="keys">Keys that should be used, keys omitted will not be included</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None, null);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items,
                                                   bool printHeaders,
                                                   TableStyles? tableStyle,
                                                   IEnumerable<string> keys)
        {
            LoadFromDictionariesParams? param = new LoadFromDictionariesParams { PrintHeaders = printHeaders, TableStyle = tableStyle };

            if (keys != null && keys.Any())
            {
                param.SetKeys(keys.ToArray());
            }

            LoadFromDictionaries? func = new LoadFromDictionaries(this, items, param);

            return func.Load();
        }
#if !NET35 && !NET40
        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries</param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="keys">Keys that should be used, keys omitted will not be included</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None, null);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<dynamic> items, bool printHeaders, TableStyles? tableStyle, IEnumerable<string> keys)
        {
            LoadFromDictionariesParams? param = new LoadFromDictionariesParams { PrintHeaders = printHeaders, TableStyle = tableStyle };

            if (keys != null && keys.Any())
            {
                param.SetKeys(keys.ToArray());
            }

            LoadFromDictionaries? func = new LoadFromDictionaries(this, items, param);

            return func.Load();
        }
#endif

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/ExpandoObjects</param>
        /// <param name="paramsConfig"><see cref="Action{LoadFromDictionariesParams}"/> to provide parameters to the function</param>
        /// <example>
        /// sheet.Cells["C1"].LoadFromDictionaries(items, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, Action<LoadFromDictionariesParams> paramsConfig)
        {
            LoadFromDictionariesParams? param = new LoadFromDictionariesParams();
            paramsConfig.Invoke(param);
            LoadFromDictionaries? func = new LoadFromDictionaries(this, items, param);

            return func.Load();
        }

#if !NET35 && !NET40
        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/ExpandoObjects</param>
        /// <param name="paramsConfig"><see cref="Action{LoadFromDictionariesParams}"/> to provide parameters to the function</param>
        /// <example>
        /// sheet.Cells["C1"].LoadFromDictionaries(items, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<dynamic> items, Action<LoadFromDictionariesParams> paramsConfig)
        {
            LoadFromDictionariesParams? param = new LoadFromDictionariesParams();
            paramsConfig.Invoke(param);
            LoadFromDictionaries? func = new LoadFromDictionaries(this, items, param);

            return func.Load();
        }
#endif

        #endregion
    }
}