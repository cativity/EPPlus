/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           Release EPPlus 5
 *************************************************************************************************/
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// Data binning properties
/// </summary>
public class ExcelChartExSerieBinning : XmlHelper
{
    internal ExcelChartExSerieBinning(XmlNamespaceManager ns, XmlNode node) :
        base(ns, node)
    {

    }
    const string _binSizePath = "cx:layoutPr/cx:binning/cx:binSize/@val";
    /// <summary>
    /// The binning by bin size. Setting this property clears the <see cref="Count"/> property
    /// </summary>
    public double? Size 
    { 
        get
        {
            return this.GetXmlNodeDoubleNull(_binSizePath);
        }
        set
        {
            this.DeleteNode(ExcelHistogramChartSerie._aggregationPath);
            this.DeleteNode(_binCountPath);
            this.SetXmlNodeDouble(_binSizePath, value);
        }
    }
    const string _binCountPath = "cx:layoutPr/cx:binning/cx:binCount/@val";
    /// <summary>
    /// The binning by bin count. Setting this property clears the <see cref="Size"/> property
    /// </summary>
    public int? Count 
    {
        get
        {
            return this.GetXmlNodeIntNull(_binCountPath);
        }
        set
        {
            this.DeleteNode(ExcelHistogramChartSerie._aggregationPath);
            this.DeleteNode(_binSizePath);
            this.SetXmlNodeInt(_binCountPath, value);
        }
    }
    const string _intervalClosedPath = "cx:layoutPr/cx:binning/@intervalClosed";
    /// <summary>
    /// The interval closed side.
    /// </summary>
    public eIntervalClosed IntervalClosed 
    { 
        get
        {
            string? v= this.GetXmlNodeString(_intervalClosedPath);
            if(v=="l")
            {
                return eIntervalClosed.Left;
            }
            if(v=="r")
            {
                return eIntervalClosed.Right;
            }
            return eIntervalClosed.None;
        }
        set
        {
            this.DeleteNode(ExcelHistogramChartSerie._aggregationPath);
            if (value==eIntervalClosed.Left)
            {
                this.SetXmlNodeString(_intervalClosedPath, "l");
            }
            else if (value == eIntervalClosed.Right)
            {
                this.SetXmlNodeString(_intervalClosedPath, "r");
            }
            else
            {
                this.DeleteNode(_intervalClosedPath);
            }
        }
    }
    const string _underflowPath = "cx:layoutPr/cx:binning/@underflow";
    /// <summary>
    /// The custom value for underflow bin is set to automatic.
    /// </summary>
    public bool UnderflowAutomatic 
    {
        get
        {
            return this.GetXmlNodeString(_underflowPath)=="auto";
        }
        set
        {
            this.DeleteNode(_intervalClosedPath);
            this.SetXmlNodeString(_underflowPath, "auto");
        }
    }
    /// <summary>
    /// A custom value for underflow bin.
    /// </summary>
    public double? Underflow 
    {
        get
        {
            return this.GetXmlNodeDoubleNull(_underflowPath);
        }
        set
        {
            this.DeleteNode(_intervalClosedPath);
            this.SetXmlNodeDouble(_underflowPath, value);
        }
    }
    const string _overflowPath = "cx:layoutPr/cx:binning/@overflow";
    /// <summary>
    /// The custom value for overflow bin is set to automatic.
    /// </summary>
    public bool OverflowAutomatic 
    {
        get
        {
            return this.GetXmlNodeString(_overflowPath) == "auto";
        }
        set
        {
            this.DeleteNode(_intervalClosedPath);
            this.SetXmlNodeString(_overflowPath, "auto");
        }
    }
    /// <summary>
    /// A custom value for overflow bin.
    /// </summary>
    public double? Overflow 
    {
        get
        {
            return this.GetXmlNodeDoubleNull(_overflowPath);
        }
        set
        {
            this.DeleteNode(_intervalClosedPath);
            this.SetXmlNodeDouble(_overflowPath, value);
        }
    }
}