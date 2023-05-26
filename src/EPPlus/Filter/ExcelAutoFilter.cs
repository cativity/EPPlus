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
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Xml;
using OfficeOpenXml.Core.CellStore;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Represents an Autofilter for a worksheet or a filter of a table
    /// </summary>
    public class ExcelAutoFilter : XmlHelper
    {
        private const string AutoFilterGuid= "71E0E44A-7884-43F4-9E11-E314B2584A5E";
        private ExcelWorksheet _worksheet;
        private ExcelTable _table;
        internal ExcelAutoFilter(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelWorksheet worksheet) : base(namespaceManager, topNode)
        {
            _columns = new ExcelFilterColumnCollection(namespaceManager, topNode, this);
            _worksheet = worksheet;
        }
        internal ExcelAutoFilter(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelTable table) : base(namespaceManager, topNode)
        {
            _columns = new ExcelFilterColumnCollection(namespaceManager, topNode, this);
            _worksheet = table.WorkSheet;
            _table = table;
        }

        internal void Save()
        {
            ApplyFilter();
            foreach (ExcelFilterColumn? c in Columns)
            {
                c.Save();
            }
        }
        /// <summary>
        /// Applies the filter, hiding rows not matching the filter columns
        /// </summary>
        /// <param name="calculateRange">If true, any formula in the autofilter range will be calculated before the filter is applied.</param>
        public void ApplyFilter(bool calculateRange=false)
        {
            if(calculateRange && _address!=null && ExcelAddressBase.IsValidAddress(_address._address))
            {
                _worksheet.Cells[_address._address].Calculate();
            }

            foreach (ExcelFilterColumn? column in Columns)
            {
                column.SetFilterValue(_worksheet, Address);
            }
            for (int row=Address._fromRow+1; row <= _address._toRow;row++)
            {
                RowInternal? rowInternal = ExcelRow.GetRowInternal(_worksheet, row);
                rowInternal.Hidden = false;
                foreach(ExcelFilterColumn? column in Columns)
                {
                    ExcelValue value = _worksheet.GetCoreValueInner(row, Address._fromCol + column.Position);
                    string? text = ValueToTextHandler.GetFormattedText(value._value, _worksheet.Workbook, value._styleId, false);
                    if (column.Match(value._value, text) == false)
                    {
                        rowInternal.Hidden = true;
                        break;
                    }
                }
            }
        }

        ExcelAddressBase _address = null;
        /// <summary>
        /// The range of the autofilter
        /// </summary>
        public ExcelAddressBase Address
        {
            get
            {
                if (_address == null)
                {
                    _address = new ExcelAddressBase(GetXmlNodeString("@ref"));
                }
                return _address;
            }
            internal set
            {                
                if (value._fromCol != Address._fromCol || value._toCol != Address._toCol || value._fromRow!=Address._fromRow) //Allow different _toRow
                {
                    _columns = new ExcelFilterColumnCollection(NameSpaceManager, TopNode, this);
                }
                SetXmlNodeString("@ref", value.Address);
                _address = value;
            }
        }

        ExcelFilterColumnCollection _columns;
        /// <summary>
        /// The columns to filter
        /// </summary>
        public ExcelFilterColumnCollection Columns
        {
            get
            {
                return _columns;
            }
        }
    }
}
