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

using OfficeOpenXml.Core;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.IO;
using System.Reflection;
using System.Xml;

namespace OfficeOpenXml.Style.Table;

/// <summary>
/// A custom named table style that applies to tables only
/// </summary>
public class ExcelTableNamedStyle : ExcelTableNamedStyleBase
{
    internal ExcelTableNamedStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles)
        : base(nameSpaceManager, topNode, styles)
    {
    }

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

    /// <summary>
    /// If the style applies to tables, pivot table or both
    /// </summary>
    public override eTableNamedStyleAppliesTo AppliesTo => eTableNamedStyleAppliesTo.Tables;
}