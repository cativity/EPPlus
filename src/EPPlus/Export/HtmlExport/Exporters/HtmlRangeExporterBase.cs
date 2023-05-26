/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  6/4/2022         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.Core.CellStore;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlRangeExporterBase : AbstractHtmlExporter
    {
        public HtmlRangeExporterBase(HtmlExportSettings settings, ExcelRangeBase range)
        {
            this.Settings = settings;
            Require.Argument(range).IsNotNull("range");
            this._ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            if (range.Addresses == null)
            {
                this.AddRange(range);
            }
            else
            {
                foreach (ExcelAddressBase? address in range.Addresses)
                {
                    this.AddRange(range.Worksheet.Cells[address.Address]);
                }
            }

            this.LoadRangeImages(this._ranges._list);
        }

        public HtmlRangeExporterBase(HtmlExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
        {
            this.Settings = settings;
            Require.Argument(ranges).IsNotNull("ranges");
            this._ranges = ranges;

            this.LoadRangeImages(this._ranges._list);
        }

        public HtmlRangeExporterBase(EPPlusReadOnlyList<ExcelRangeBase> ranges)
        {
            Require.Argument(ranges).IsNotNull("ranges");
            this._ranges = ranges;

            this.LoadRangeImages(this._ranges._list);
        }

        protected List<int> _columns = new List<int>();
        protected HtmlExportSettings Settings;
        protected readonly List<ExcelAddressBase> _mergedCells = new List<ExcelAddressBase>();

        protected void LoadVisibleColumns(ExcelRangeBase range)
        {
            ExcelWorksheet? ws = range.Worksheet;
            this._columns = new List<int>();
            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                ExcelColumn? c = ws.GetColumn(col);
                if (c == null || (c.Hidden == false && c.Width > 0))
                {
                    this._columns.Add(col);
                }
            }
        }

        protected EPPlusReadOnlyList<ExcelRangeBase> _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

        private void AddRange(ExcelRangeBase range)
        {
            if (range.IsFullColumn && range.IsFullRow)
            {
                this._ranges.Add(new ExcelRangeBase(range.Worksheet, range.Worksheet.Dimension.Address));
            }
            else
            {
                this._ranges.Add(range);
            }
        }

        protected void ValidateRangeIndex(int rangeIndex)
        {
            if (rangeIndex < 0 || rangeIndex >= this._ranges.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(rangeIndex));
            }
        }

        internal static bool HandleHiddenRow(EpplusHtmlWriter writer, ExcelWorksheet ws, HtmlExportSettings Settings, ref int row)
        {
            if (Settings.HiddenRows != eHiddenState.Include)
            {
                ExcelRow? r = ws.Row(row);
                if (r.Hidden || r.Height == 0)
                {
                    if (Settings.HiddenRows == eHiddenState.IncludeButHide)
                    {
                        writer.AddAttribute("class", $"{Settings.StyleClassPrefix}hidden");
                    }
                    else
                    {
                        row++;
                        return true;
                    }
                }
            }

            return false;
        }

        internal static void AddRowHeightStyle(EpplusHtmlWriter writer, ExcelRangeBase range, int row, string styleClassPrefix, bool isMultiSheet)
        {
            ExcelValue r = range.Worksheet._values.GetValue(row, 0);
            if (r._value is RowInternal rowInternal)
            {
                if (rowInternal.Height != -1 && rowInternal.Height != range.Worksheet.DefaultRowHeight)
                {
                    writer.AddAttribute("style", $"height:{rowInternal.Height}pt");
                    return;
                }
            }

            string? clsName = HtmlExportTableUtil.GetWorksheetClassName(styleClassPrefix, "drh", range.Worksheet, isMultiSheet);
            writer.AddAttribute("class", clsName); //Default row height
        }

        protected static string GetPictureName(HtmlImage p)
        {
            string? hash = ((IPictureContainer)p.Picture).ImageHash;
            FileInfo? fi = new FileInfo(p.Picture.Part.Uri.OriginalString);
            string? name = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length);

            return HtmlExportTableUtil.GetClassName(name, hash);
        }

        protected bool InMergeCellSpan(int row, int col)
        {
            for (int i = 0; i < this._mergedCells.Count; i++)
            {
                ExcelAddressBase? adr = this._mergedCells[i];
                if (adr._toRow < row || (adr._toRow == row && adr._toCol < col))
                {
                    this._mergedCells.RemoveAt(i);
                    i--;
                }
                else
                {
                    if (row >= adr._fromRow && row <= adr._toRow &&
                       col >= adr._fromCol && col <= adr._toCol)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        protected void SetColRowSpan(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelRange cell)
        {
            if (cell.Merge)
            {
                string? address = cell.Worksheet.MergedCells[cell._fromRow, cell._fromCol];
                if (address != null)
                {
                    ExcelAddressBase? ma = new ExcelAddressBase(address);
                    bool added = false;
                    //ColSpan
                    if (ma._fromCol == cell._fromCol || range._fromCol == cell._fromCol)
                    {
                        int maxCol = Math.Min(ma._toCol, range._toCol);
                        int colSpan = maxCol - ma._fromCol + 1;
                        if (colSpan > 1)
                        {
                            writer.AddAttribute("colspan", colSpan.ToString(CultureInfo.InvariantCulture));
                        }

                        this._mergedCells.Add(ma);
                        added = true;
                    }
                    //RowSpan
                    if (ma._fromRow == cell._fromRow || range._fromRow == cell._fromRow)
                    {
                        int maxRow = Math.Min(ma._toRow, range._toRow);
                        int rowSpan = maxRow - ma._fromRow + 1;
                        if (rowSpan > 1)
                        {
                            writer.AddAttribute("rowspan", rowSpan.ToString(CultureInfo.InvariantCulture));
                        }
                        if (added == false)
                        {
                            this._mergedCells.Add(ma);
                        }
                    }
                }
            }
        }

        protected void GetDataTypes(ExcelRangeBase range, HtmlRangeExportSettings settings)
        {
            if (range._fromRow + settings.HeaderRows > ExcelPackage.MaxRows)
            {
                throw new InvalidOperationException("Range From Row + Header rows is out of bounds");
            }

            this._dataTypes = new List<string>();
            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                this._dataTypes.Add(
                                    ColumnDataTypeManager.GetColumnDataType(range.Worksheet, range, range._fromRow + settings.HeaderRows, col));
            }
        }
        bool? _isMultiSheet = null;
        protected bool IsMultiSheet
        {
            get
            {
                if (this._isMultiSheet.HasValue == false)
                {
                    this._isMultiSheet = this._ranges.Select(x => x.Worksheet).Distinct().Count() > 1;
                }
                return this._isMultiSheet.Value;
            }
        }

        protected static void AddTableAccessibilityAttributes(AccessibilitySettings settings, EpplusHtmlWriter writer)
        {
            if (!settings.TableSettings.AddAccessibilityAttributes)
            {
                return;
            }

            if (!string.IsNullOrEmpty(settings.TableSettings.TableRole))
            {
                writer.AddAttribute("role", settings.TableSettings.TableRole);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaLabel))
            {
                writer.AddAttribute(AriaAttributes.AriaLabel.AttributeName, settings.TableSettings.AriaLabel);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaLabelledBy))
            {
                writer.AddAttribute(AriaAttributes.AriaLabelledBy.AttributeName, settings.TableSettings.AriaLabelledBy);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaDescribedBy))
            {
                writer.AddAttribute(AriaAttributes.AriaDescribedBy.AttributeName, settings.TableSettings.AriaDescribedBy);
            }
        }

        protected string GetTableId(int index, ExcelHtmlOverrideExportSettings overrideSettings)
        {
            if (overrideSettings == null || string.IsNullOrEmpty(overrideSettings.TableId))
            {
                if (this._ranges.Count > 1 && !string.IsNullOrEmpty(this.Settings.TableId))
                {
                    return this.Settings.TableId + index.ToString(CultureInfo.InvariantCulture);
                }
                return this.Settings.TableId;
            }
            return overrideSettings.TableId;
        }

        protected List<string> GetAdditionalClassNames(ExcelHtmlOverrideExportSettings overrideSettings)
        {
            if (overrideSettings == null || overrideSettings.AdditionalTableClassNames == null)
            {
                return this.Settings.AdditionalTableClassNames;
            }

            return overrideSettings.AdditionalTableClassNames;
        }

        protected AccessibilitySettings GetAccessibilitySettings(ExcelHtmlOverrideExportSettings overrideSettings)
        {
            if (overrideSettings == null || overrideSettings.Accessibility == null)
            {
                return this.Settings.Accessibility;
            }

            return overrideSettings.Accessibility;
        }

        protected static void AddClassesAttributes(EpplusHtmlWriter writer, ExcelTable table, string tableId, List<string> additionalTableClassNames)
        {
            string? tableClasses = TableClass;
            if (table != null)
            {
                tableClasses += " " + HtmlExportTableUtil.GetTableClasses(table); //Add classes for the table styles if the range corresponds to a table.
            }
            if (additionalTableClassNames != null && additionalTableClassNames.Count > 0)
            {
                foreach (string? cls in additionalTableClassNames)
                {
                    tableClasses += $" {cls}";
                }
            }
            writer.AddAttribute(HtmlAttributes.Class, $"{tableClasses}");

            if (!string.IsNullOrEmpty(tableId))
            {
                writer.AddAttribute(HtmlAttributes.Id, tableId);
            }
        }
    }
}
