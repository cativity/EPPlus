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
using OfficeOpenXml.Style;
using System.Data;
namespace OfficeOpenXml
{
    /// <summary>
    /// A range of cells. 
    /// </summary>
    public class ExcelRange : ExcelRangeBase
    {
        #region "Constructors"
        internal ExcelRange(ExcelWorksheet sheet, string address)
            : base(sheet, address)
        {

        }
        internal ExcelRange(ExcelWorksheet sheet, int fromRow, int fromCol, int toRow, int toCol)
            : base(sheet)
        {
            this._fromRow = fromRow;
            this._fromCol = fromCol;
            this._toRow = toRow;
            this._toCol = toCol;
        }
        #endregion
        #region "Indexers"
        /// <summary>
        /// Access the range using an address
        /// </summary>
        /// <param name="Address">The address</param>
        /// <returns>A range object</returns>
        public ExcelRange this[string Address]
        {
            get
            {
                if (this._worksheet.Names.ContainsKey(Address))
                {
                    if (this._worksheet.Names[Address].IsName)
                    {
                        return null;
                    }
                    else
                    {
                        this.Address = this._worksheet.Names[Address].Address;
                    }
                }
                else
                {
                    if(Address.IndexOfAny(new char[] { '\'', '[', '!' })>=0)
                    {
                        ExcelAddress? a = new ExcelAddress(Address);
                        if(a.WorkSheetName!=null && a.WorkSheetName.Equals(this._worksheet.Name, StringComparison.InvariantCultureIgnoreCase)==false)
                        {
                            throw new InvalidOperationException($"The worksheet address {Address} is not within the worksheet {this._worksheet.Name}");
                        }
                    }

                    this.SetAddress(Address, this._workbook, this._worksheet.Name);
                    this.ChangeAddress();
                }
                if((this._fromRow < 1 || this._fromCol < 1) && Address.Equals("#REF!", StringComparison.InvariantCultureIgnoreCase)==false)
                {
                    throw (new InvalidOperationException("Address is not valid."));
                }

                this._rtc = null;
                return this;
            }
        }

        private static ExcelRange GetTableAddess(ExcelWorksheet _worksheet, string address)
        {
            int ixStart = address.IndexOf('[');
            if (ixStart == 0) //External Address
            {
                int ixEnd = address.IndexOf(']',ixStart+1);
                if (ixStart >= 0 & ixEnd >= 0)
                {
                    string? external = address.Substring(ixStart + 1, ixEnd - 1);
                    //if (Worksheet.Workbook._externalReferences.Count < external)
                    //{
                    //foreach(var 
                    //}
                }
            }
            return null;
        }
        /// <summary>
        /// Access a single cell
        /// </summary>
        /// <param name="Row">The row</param>
        /// <param name="Col">The column</param>
        /// <returns>A range object</returns>
        public ExcelRange this[int Row, int Col]
        {
            get
            {
                ValidateRowCol(Row, Col);

                this._fromCol = Col;
                this._fromRow = Row;
                this._toCol = Col;
                this._toRow = Row;
                this._rtc = null;
                // avoid address re-calculation
                //base.Address = GetAddress(_fromRow, _fromCol);
                this._start = null;
                this._end = null;
                this._addresses = null;
                this._address = GetAddress(this._fromRow, this._fromCol);
                this.ChangeAddress();
                return this;
            }
        }
        /// <summary>
        /// Access a range of cells
        /// </summary>
        /// <param name="FromRow">Start row</param>
        /// <param name="FromCol">Start column</param>
        /// <param name="ToRow">End Row</param>
        /// <param name="ToCol">End Column</param>
        /// <returns></returns>
        public ExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol]
        {
            get
            {
                ValidateRowCol(FromRow, FromCol);
                ValidateRowCol(ToRow, ToCol);

                this._fromCol = FromCol;
                this._fromRow = FromRow;
                this._toCol = ToCol;
                this._toRow = ToRow;
                this._rtc = null;
                // avoid address re-calculation
                //base.Address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
                this._start = null;
                this._end = null;
                this._addresses = null;
                this._address = GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol);
                this.ChangeAddress();
                return this;
            }
        }
        #endregion
        private static void ValidateRowCol(int Row, int Col)
        {
            if (Row < 1 || Row > ExcelPackage.MaxRows)
            {
                throw (new ArgumentException("Row out of range"));
            }
            if (Col < 1 || Col > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentException("Column out of range"));
            }
        }

    }
}
