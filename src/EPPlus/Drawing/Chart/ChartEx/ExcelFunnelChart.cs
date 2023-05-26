﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// Represents a Funnel Chart
/// </summary>
public class ExcelFunnelChart : ExcelChartEx
{
    internal ExcelFunnelChart(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, XmlDocument chartXml = null, ExcelGroupShape parent = null) :
        base(drawings, drawingsNode, type, chartXml, parent)
    {
        this.StyleManager.SetChartStyle(Chart.Style.ePresetChartStyle.FunnelChartStyle1);
    }
    internal ExcelFunnelChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
        base(drawings, node, uriChart, part, chartXml, chartNode, parent)
    {
    }
}