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
namespace OfficeOpenXml.Style;

internal class StyleChangeEventArgs : EventArgs
{
    internal StyleChangeEventArgs(eStyleClass styleclass, eStyleProperty styleProperty, object value, int positionID, string address)
    {
        this.StyleClass = styleclass;
        this.StyleProperty=styleProperty;
        this.Value = value;
        this.Address = address;
        this.PositionID = positionID;
    }
    internal eStyleClass StyleClass;
    internal eStyleProperty StyleProperty;
    //internal string PropertyName;
    internal object Value;
    internal int PositionID { get; set; }
    //internal string Address;
    internal string Address;
}