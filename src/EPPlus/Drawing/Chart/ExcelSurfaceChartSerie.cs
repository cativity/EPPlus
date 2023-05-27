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
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A serie for a surface chart
/// </summary>
public sealed class ExcelSurfaceChartSerie : ExcelChartStandardSerie
{
    /// <summary>
    /// Default constructor
    /// </summary>
    /// <param name="chart">The chart</param>
    /// <param name="ns">Namespacemanager</param>
    /// <param name="node">Topnode</param>
    /// <param name="isPivot">Is pivotchart</param>
    internal ExcelSurfaceChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
        : base(chart, ns, node, isPivot)
    {
    }
}