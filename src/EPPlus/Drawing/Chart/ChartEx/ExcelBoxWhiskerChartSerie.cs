﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/

using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Utils.Extensions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// A series for an Box &amp; Whisker Chart
/// </summary>
public class ExcelBoxWhiskerChartSerie : ExcelChartExSerie
{
    const string _path = "cx:layoutPr/cx:visibility";

    internal ExcelBoxWhiskerChartSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node)
        : base(chart, ns, node)
    {
    }

    /// <summary>
    /// The layout type for the parent labels
    /// </summary>
    public eParentLabelLayout ParentLabelLayout
    {
        get => this.GetXmlNodeString("cx:layoutPr/cx:parentLabelLayout/@val").ToEnum(eParentLabelLayout.None);
        set => this.SetXmlNodeString("cx:layoutPr/cx:parentLabelLayout/@val", value.ToEnumString());
    }

    /// <summary>
    /// The quartile calculation methods
    /// </summary>
    public eQuartileMethod? QuartileMethod
    {
        get
        {
            string? s = this.GetXmlNodeString("cx:layoutPr/cx:statistics/@quartileMethod");

            if (string.IsNullOrEmpty(s))
            {
                return null;
            }

            return s.ToEnum(eQuartileMethod.Inclusive);
        }
        set => this.SetXmlNodeString("cx:layoutPr/cx:statistics/@quartileMethod", value.ToEnumString());
    }

    /// <summary>
    /// The visibility of connector lines between data points
    /// </summary>
    public bool ShowMeanLine
    {
        get => this.GetXmlNodeBool($"{_path}/@meanLine");
        set => this.SetXmlNodeBool($"{_path}/@meanLine", value);
    }

    /// <summary>
    /// The visibility of markers denoting the mean
    /// </summary>
    public bool ShowMeanMarker
    {
        get => this.GetXmlNodeBool($"{_path}/@meanMarker");
        set => this.SetXmlNodeBool($"{_path}/@meanMarker", value);
    }

    /// <summary>
    /// The visibility of non-outlier data points
    /// </summary>
    public bool ShowNonOutliers
    {
        get => this.GetXmlNodeBool($"{_path}/@nonoutliers");
        set => this.SetXmlNodeBool($"{_path}/@nonoutliers", value);
    }

    /// <summary>
    /// The visibility of outlier data points
    /// </summary>
    public bool ShowOutliers
    {
        get => this.GetXmlNodeBool($"{_path}/@outliers");
        set => this.SetXmlNodeBool($"{_path}/@outliers", value);
    }
}