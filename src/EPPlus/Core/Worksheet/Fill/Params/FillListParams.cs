﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    08/11/2021         EPPlus Software AB       EPPlus 5.8
 *************************************************************************************************/

namespace OfficeOpenXml.Core.Worksheet.Fill;

/// <summary>
/// Parameters for the <see cref="ExcelRangeBase.FillList{T}(System.Collections.Generic.IEnumerable{T}, System.Action{FillListParams})" /> method 
/// </summary>
public class FillListParams : FillParams
{
    /// <summary>
    /// The start index in the list. 
    /// <seealso cref="FillParams.Direction"/>
    /// </summary>
    public int StartIndex { get; set; }
}