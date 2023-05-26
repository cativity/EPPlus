﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    04/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// The quartile calculation methods
/// </summary>
public enum eQuartileMethod
{
    /// <summary>
    /// The quartile calculation includes the median when splitting the dataset into quartiles
    /// </summary>
    Inclusive,
    /// <summary>
    /// The quartile calculation excludes the median when splitting the dataset into quartiles
    /// </summary>
    Exclusive
}