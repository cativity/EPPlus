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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Core;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    /// <summary>
    /// Base class for Html exporters
    /// </summary>
    public abstract class ExcelHtmlExporterBase
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="range"></param>
        internal ExcelHtmlExporterBase(ExcelRangeBase range)
        {
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

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ranges"></param>
        internal ExcelHtmlExporterBase(params ExcelRangeBase[] ranges)
        {
            this._ranges = new EPPlusReadOnlyList<ExcelRangeBase>();
            foreach (ExcelRangeBase? range in ranges)
            {
                this.AddRange(range);
            }
        }


        private readonly EPPlusReadOnlyList<ExcelRangeBase> _ranges;

        /// <summary>
        /// Exported ranges
        /// </summary>
        public EPPlusReadOnlyList<ExcelRangeBase> Ranges
        {
            get
            {
                return this._ranges;
            }
        }

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
