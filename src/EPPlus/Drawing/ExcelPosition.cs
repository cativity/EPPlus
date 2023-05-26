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
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Position of the a drawing.
    /// </summary>
    public class ExcelPosition : XmlHelper
    {
        internal delegate void SetWidthCallback();
        XmlNode _node;
        XmlNamespaceManager _ns;
        SetWidthCallback _setWidthCallback;
        internal ExcelPosition(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback) :
            base(ns, node)
        {
            this._node = node;
            this._ns = ns;
            this._setWidthCallback = setWidthCallback;
            this.Load();
        }
        const string colPath = "xdr:col";
        int _column, _row, _columnOff, _rowOff;        
        /// <summary>
        /// The column
        /// </summary>
        public int Column
        {
            get
            {
                return this._column;
            }
            set
            {
                this._column = value;
                this._setWidthCallback?.Invoke();
            }
        }
        const string rowPath = "xdr:row";
        /// <summary>
        /// The row
        /// </summary>
        public int Row
        {
            get
            {
                return this._row;
            }
            set
            {
                this._row = value;
                this._setWidthCallback?.Invoke();
            }
        }
        const string colOffPath = "xdr:colOff";
        /// <summary>
        /// Column Offset in EMU
        /// ss
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public int ColumnOff
        {
            get
            {
                return this._columnOff;
            }
            set
            {
                this._columnOff = value;
                this._setWidthCallback?.Invoke();
            }
        }
        const string rowOffPath = "xdr:rowOff";
        /// <summary>
        /// Row Offset in EMU
        /// 
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public int RowOff
        {
            get
            {
                return this._rowOff;
            }
            set
            {
                this._rowOff = value;
                this._setWidthCallback?.Invoke();
            }
        }
        public void Load()
        {
            this._column = this.GetXmlNodeInt(colPath);
            this._columnOff = this.GetXmlNodeInt(colOffPath);
            this._row = this.GetXmlNodeInt(rowPath);
            this._rowOff = this.GetXmlNodeInt(rowOffPath);
        }
        public void UpdateXml()
        {
            this.SetXmlNodeString(colPath, this._column.ToString());
            this.SetXmlNodeString(colOffPath, this._columnOff.ToString());
            this.SetXmlNodeString(rowPath, this._row.ToString());
            this.SetXmlNodeString(rowOffPath, this._rowOff.ToString());
        }
    }
}