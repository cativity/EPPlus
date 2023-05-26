﻿/*************************************************************************************************
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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class CssRangeExporterBase : AbstractHtmlExporter
    {
        public CssRangeExporterBase(HtmlExportSettings settings, ExcelRangeBase range)
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
        }

        public CssRangeExporterBase(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
        {
            this.Settings = settings;
            Require.Argument(ranges).IsNotNull("ranges");
            this._ranges = ranges;
        }

        protected HtmlExportSettings Settings;
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
    }
}
