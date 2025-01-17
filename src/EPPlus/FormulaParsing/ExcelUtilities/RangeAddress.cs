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
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class RangeAddress
{
    public RangeAddress() => this.Address = string.Empty;

    internal string Address { get; set; }

    public string Worksheet { get; internal set; }

    public int FromCol { get; internal set; }

    public int ToCol { get; internal set; }

    public int FromRow { get; internal set; }

    public int ToRow { get; internal set; }

    public override string ToString() => this.Address;

    private static RangeAddress _empty = new RangeAddress();

    public static RangeAddress Empty => _empty;

    /// <summary>
    /// Returns true if this range collides (full or partly) with the supplied range
    /// </summary>
    /// <param name="other">The range to check</param>
    /// <returns></returns>
    public bool CollidesWith(RangeAddress other)
    {
        if (other.Worksheet != this.Worksheet)
        {
            return false;
        }

        if (other.FromRow > this.ToRow || other.FromCol > this.ToCol || this.FromRow > other.ToRow || this.FromCol > other.ToCol)
        {
            return false;
        }

        return true;
    }
}