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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class CssRangeExporterAsync : CssRangeExporterBase
    {
        public CssRangeExporterAsync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
            : base(settings, ranges) =>
            this._settings = settings;

        public CssRangeExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase range)
            : base(settings, range) =>
            this._settings = settings;

        private readonly HtmlRangeExportSettings _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetCssStringAsync()
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            await this.RenderCssAsync(ms);
            ms.Position = 0;
            using StreamReader? sr = new StreamReader(ms);

            return await sr.ReadToEndAsync();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to html and writes it to a stream
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns></returns>
        public async Task RenderCssAsync(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writable System.IO.Stream");
            }

            //if (_datatypes.Count == 0) GetDataTypes();
            StreamWriter? sw = new StreamWriter(stream);
            await this.RenderCellCssAsync(sw);
        }

        private async Task RenderCellCssAsync(StreamWriter sw)
        {
            EpplusCssWriter? styleWriter =
                new EpplusCssWriter(sw, this._ranges._list, this._settings, this._settings.Css, this._settings.Css.CssExclude, this._styleCache);

            await styleWriter.RenderAdditionalAndFontCssAsync(TableClass);
            HashSet<TableStyles>? addedTableStyles = new HashSet<TableStyles>();

            foreach (ExcelRangeBase? range in this._ranges._list)
            {
                ExcelWorksheet? ws = range.Worksheet;
                ExcelStyles? styles = ws.Workbook.Styles;

                CellStoreEnumerator<ExcelValue>? ce =
                    new CellStoreEnumerator<ExcelValue>(range.Worksheet._values, range._fromRow, range._fromCol, range._toRow, range._toCol);

                ExcelAddressBase address = null;

                while (ce.Next())
                {
                    if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
                    {
                        string? ma = ws.MergedCells[ce.Row, ce.Column];

                        if (ma != null)
                        {
                            if (address == null || address.Address != ma)
                            {
                                address = new ExcelAddressBase(ma);
                            }

                            int fromRow = address._fromRow < range._fromRow ? range._fromRow : address._fromRow;
                            int fromCol = address._fromCol < range._fromCol ? range._fromCol : address._fromCol;

                            if (fromRow != ce.Row || fromCol != ce.Column) //Only add the style for the top-left cell in the merged range.
                            {
                                continue;
                            }

                            ExcelAddressBase? mAdr = new ExcelAddressBase(ma);
                            int bottomStyleId = range.Worksheet._values.GetValue(mAdr._toRow, mAdr._fromCol)._styleId;
                            int rightStyleId = range.Worksheet._values.GetValue(mAdr._fromRow, mAdr._toCol)._styleId;

                            await styleWriter.AddToCssAsync(styles,
                                                            ce.Value._styleId,
                                                            bottomStyleId,
                                                            rightStyleId,
                                                            this.Settings.StyleClassPrefix,
                                                            this.Settings.CellStyleClassName);
                        }
                        else
                        {
                            await styleWriter.AddToCssAsync(styles, ce.Value._styleId, this.Settings.StyleClassPrefix, this.Settings.CellStyleClassName);
                        }
                    }
                }

                if (this.Settings.TableStyle == eHtmlRangeTableInclude.Include)
                {
                    ExcelTable? table = range.GetTable();

                    if (table != null && table.TableStyle != TableStyles.None && addedTableStyles.Contains(table.TableStyle) == false)
                    {
                        HtmlTableExportSettings? settings = new HtmlTableExportSettings() { Minify = this.Settings.Minify };
                        await HtmlExportTableUtil.RenderTableCssAsync(sw, table, settings, this._styleCache, this._dataTypes);
                        _ = addedTableStyles.Add(table.TableStyle);
                    }
                }
            }

            if (this.Settings.Pictures.Include == ePictureInclude.Include)
            {
                this.LoadRangeImages(this._ranges._list);

                foreach (HtmlImage? p in this._rangePictures)
                {
                    await styleWriter.AddPictureToCssAsync(p);
                }
            }

            await styleWriter.FlushStreamAsync();
        }
    }
}
#endif