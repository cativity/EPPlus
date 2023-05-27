﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/29/2020         EPPlus Software AB       EPPlus 5.3
 *************************************************************************************************/

namespace OfficeOpenXml;

/// <summary>
/// The source of the slicer data
/// </summary>
public enum eSlicerSourceType
{
    /// <summary>
    /// A pivot table
    /// </summary>
    PivotTable,

    /// <summary>
    /// A table
    /// </summary>
    Table
}