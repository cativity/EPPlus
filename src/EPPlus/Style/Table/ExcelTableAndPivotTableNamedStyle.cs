﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/

using System.Xml;

namespace OfficeOpenXml.Style.Table;

/// <summary>
/// A custom named table style that applies to both tables and pivot tables
/// </summary>
public class ExcelTableAndPivotTableNamedStyle : ExcelPivotTableNamedStyle
{
    internal ExcelTableAndPivotTableNamedStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles)
        : base(nameSpaceManager, topNode, styles)
    {
    }

    /// <summary>
    /// If the style applies to tables, pivot table or both
    /// </summary>
    public override eTableNamedStyleAppliesTo AppliesTo => eTableNamedStyleAppliesTo.TablesAndPivotTables;

    /// <summary>
    /// Applies to the last header cell of a table
    /// </summary>
    public ExcelTableStyleElement LastHeaderCell => this.GetTableStyleElement(eTableStyleElement.LastHeaderCell);

    /// <summary>
    /// Applies to the first total cell of a table
    /// </summary>
    public ExcelTableStyleElement FirstTotalCell => this.GetTableStyleElement(eTableStyleElement.FirstTotalCell);

    /// <summary>
    /// Applies to the last total cell of a table
    /// </summary>
    public ExcelTableStyleElement LastTotalCell => this.GetTableStyleElement(eTableStyleElement.LastTotalCell);
}