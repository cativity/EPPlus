﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Exports a <see cref="ExcelTable"/> to Html
    /// </summary>
    public partial class TableExporter
    {
        internal TableExporter(ExcelTable table)
        {
            Require.Argument(table).IsNotNull("table");
            _table = table;
        }

        private readonly ExcelTable _table;
        private const string TableClass = "epplus-table";
        private const string TableStyleClassPrefix = "epplus-tablestyle-";
        private readonly CellDataWriter _cellDataWriter = new CellDataWriter();

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            return GetHtmlString(HtmlTableExportOptions.Default);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="options"><see cref="HtmlTableExportOptions">Options</see> for the export</param>
        /// <returns>A html table</returns>
        public string GetHtmlString(HtmlTableExportOptions options)
        {
            using(var ms = new MemoryStream())
            {
                RenderHtml(ms, options);
                ms.Position = 0;
                using(var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        public void RenderHtml(Stream stream)
        {
            RenderHtml(stream, HtmlTableExportOptions.Default);
        }

        public void RenderHtml(Stream stream, HtmlTableExportOptions options)
        {
            Require.Argument(options).IsNotNull("options");
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            var writer = new EpplusHtmlWriter(stream);
            if (_table.TableStyle != TableStyles.None)
            {
                writer.AddAttribute(HtmlAttributes.Class, $"{TableClass} {TableStyleClassPrefix}{_table.TableStyle.ToString().ToLowerInvariant()}");
            }
            else
            {
                writer.AddAttribute(HtmlAttributes.Class, $"{TableClass}");
            }
            if(!string.IsNullOrEmpty(options.TableId))
            {
                writer.AddAttribute(HtmlAttributes.Id, options.TableId);
            }
            writer.RenderBeginTag(HtmlElements.Table);

            writer.ApplyFormatIncreaseIndent(options.FormatHtml);
            if (_table.ShowHeader)
            {
                RenderHeaderRow(writer, options);
            }
            // table rows
            RenderTableRows(writer, options);
            // end tag table
            writer.RenderEndTag();

        }

        private void RenderTableRows(EpplusHtmlWriter writer, HtmlTableExportOptions options)
        {
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(options.FormatHtml);
            var rowIndex = _table.ShowTotal ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            while (rowIndex < _table.Address._toRow)
            {
                rowIndex++;
                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(options.FormatHtml);
                var tableRange = _table.WorkSheet.Cells[rowIndex, _table.Address._fromCol, rowIndex, _table.Address._toCol];
                foreach (var cell in tableRange)
                {
                    _cellDataWriter.Write(cell, writer, options);
                }
                // end tag tr
                writer.Indent--;
                writer.RenderEndTag();
            }

            writer.ApplyFormatDecreaseIndent(options.FormatHtml);
            // end tag tbody
            writer.RenderEndTag();
            writer.ApplyFormatDecreaseIndent(options.FormatHtml);
        }

        private void RenderHeaderRow(EpplusHtmlWriter writer, HtmlTableExportOptions options)
        {
            // table header row
            var rowIndex = _table.Address._fromRow;
            writer.RenderBeginTag(HtmlElements.Thead);
            writer.ApplyFormatIncreaseIndent(options.FormatHtml);
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(options.FormatHtml);
            var headerRange = _table.WorkSheet.Cells[rowIndex, _table.Address._fromCol, rowIndex, _table.Address._toCol];
            var col = 1;
            foreach (var cell in headerRange)
            {
                var dataType = ColumnDataTypeManager.GetColumnDataType(_table.WorkSheet, _table.Range, 2, col++);
                writer.AddAttribute("data-datatype", dataType);
                writer.RenderBeginTag(HtmlElements.TableHeader);
                // TODO: apply format
                writer.Write(cell.Value.ToString());
                writer.RenderEndTag();
                writer.ApplyFormat(options.FormatHtml);

            }
            writer.Indent--;
            writer.RenderEndTag();
            writer.ApplyFormatDecreaseIndent(options.FormatHtml);
            writer.RenderEndTag();
            writer.ApplyFormat(options.FormatHtml);
        }


    }
}
