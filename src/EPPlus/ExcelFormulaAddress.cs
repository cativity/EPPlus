﻿/*************************************************************************************************
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

using System.Collections.Generic;

namespace OfficeOpenXml;

/// <summary>
/// Range address used in the formula parser
/// </summary>
public class ExcelFormulaAddress : ExcelAddressBase
{
    /// <summary>
    /// Creates a Address object
    /// </summary>
    internal ExcelFormulaAddress()
        : base()
    {
    }

    /// <summary>
    /// Creates an Address object
    /// </summary>
    /// <param name="fromRow">start row</param>
    /// <param name="fromCol">start column</param>
    /// <param name="toRow">End row</param>
    /// <param name="toColumn">End column</param>
    public ExcelFormulaAddress(int fromRow, int fromCol, int toRow, int toColumn)
        : base(fromRow, fromCol, toRow, toColumn) =>
        this._ws = "";

    /// <summary>
    /// Creates an Address object
    /// </summary>
    /// <param name="address">The formula address</param>
    /// <param name="worksheet">The worksheet</param>
    public ExcelFormulaAddress(string address, ExcelWorksheet worksheet)
        : base(address, worksheet?.Workbook, worksheet?.Name) =>
        this.SetFixed();

    internal ExcelFormulaAddress(string ws, string address)
        : base(address)
    {
        if (string.IsNullOrEmpty(this._ws))
        {
            this._ws = ws;
        }

        this.SetFixed();
    }

    internal ExcelFormulaAddress(string ws, string address, bool isName)
        : base(address, isName)
    {
        if (string.IsNullOrEmpty(this._ws))
        {
            this._ws = ws;
        }

        if (!isName)
        {
            this.SetFixed();
        }
    }

    private void SetFixed()
    {
        if (this.Address.IndexOf('[') >= 0)
        {
            return;
        }

        string? address = this.FirstAddress;

        if (this._fromRow == this._toRow && this._fromCol == this._toCol)
        {
            GetFixed(address, out this._fromRowFixed, out this._fromColFixed);
        }
        else
        {
            string[]? cells = address.Split(':');

            if (cells.Length > 1) //If 1 then the address is the entire worksheet
            {
                GetFixed(cells[0], out this._fromRowFixed, out this._fromColFixed);
                GetFixed(cells[1], out this._toRowFixed, out this._toColFixed);
            }
        }
    }

    private static void GetFixed(string address, out bool rowFixed, out bool colFixed)
    {
        rowFixed = colFixed = false;
        int ix = address.IndexOf('$');

        while (ix > -1)
        {
            ix++;

            if (ix < address.Length)
            {
                if (address[ix] >= '0' && address[ix] <= '9')
                {
                    rowFixed = true;

                    break;
                }
                else
                {
                    colFixed = true;
                }
            }

            ix = address.IndexOf('$', ix);
        }
    }

    /// <summary>
    /// The address for the range
    /// </summary>
    /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
    public new string Address
    {
        get
        {
            if (string.IsNullOrEmpty(this._address) && this._fromRow > 0)
            {
                this._address = GetAddress(this._fromRow,
                                           this._fromCol,
                                           this._toRow,
                                           this._toCol,
                                           this._fromRowFixed,
                                           this._toRowFixed,
                                           this._fromColFixed,
                                           this._toColFixed);
            }

            return this._address;
        }
        set
        {
            this.SetAddress(value, null, null);
            this.ChangeAddress();
            this.SetFixed();
        }
    }

    internal new List<ExcelFormulaAddress> _addresses;

    /// <summary>
    /// Addresses can be separated by a comma. If the address contains multiple addresses this list contains them.
    /// </summary>
    public new List<ExcelFormulaAddress> Addresses => this._addresses ??= new List<ExcelFormulaAddress>();

    internal string GetOffset(int row, int column, bool withWbWs = false)
    {
        int fromRow = this._fromRow,
            fromCol = this._fromCol,
            toRow = this._toRow,
            tocol = this._toCol;

        bool isMulti = fromRow != toRow || fromCol != tocol;

        if (!this._fromRowFixed)
        {
            fromRow += row;
        }

        if (!this._fromColFixed)
        {
            fromCol += column;
        }

        if (isMulti)
        {
            if (!this._toRowFixed)
            {
                toRow += row;
            }

            if (!this._toColFixed)
            {
                tocol += column;
            }
        }
        else
        {
            toRow = fromRow;
            tocol = fromCol;
        }

        string a = GetAddress(fromRow, fromCol, toRow, tocol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed);

        if (this.Addresses != null)
        {
            foreach (ExcelFormulaAddress? sa in this.Addresses)
            {
                a += "," + sa.GetOffset(row, column, withWbWs);
            }
        }

        if (withWbWs)
        {
            return this.GetAddressWorkBookWorkSheet() + a;
        }
        else
        {
            return a;
        }
    }
}