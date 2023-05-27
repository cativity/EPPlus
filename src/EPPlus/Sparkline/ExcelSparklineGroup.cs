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

using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sparkline;

/// <summary>
/// Represents a group of sparklines
/// </summary>
public class ExcelSparklineGroup : XmlHelper
{
    // Schema here... https://msdn.microsoft.com/en-us/library/hh656506(v=office.12).aspx

    ExcelWorksheet _ws;

    internal ExcelSparklineGroup(XmlNamespaceManager ns, XmlElement topNode, ExcelWorksheet ws)
        : base(ns, topNode)
    {
        this.SchemaNodeOrder = new string[]
        {
            "colorSeries", "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh", "colorLow", "f", "sparklines"
        };

        this.Sparklines = new ExcelSparklineCollection(this);
        this._ws = ws;
    }

    /// <summary>
    /// The range containing the dateaxis from the sparklines.
    /// Set to Null to remove the dateaxis.
    /// </summary>
    public ExcelRangeBase DateAxisRange
    {
        get
        {
            string? f = this.GetXmlNodeString("xm:f");

            if (string.IsNullOrEmpty(f))
            {
                return null;
            }

            ExcelAddressBase? a = new ExcelAddressBase(f);

            if (a.WorkSheetName.Equals(this._ws.Name, StringComparison.CurrentCultureIgnoreCase))
            {
                return this._ws.Cells[a.Address];
            }
            else
            {
                ExcelWorksheet? ws = this._ws.Workbook.Worksheets[a.WorkSheetName];

                return ws.Cells[a.Address];
            }
        }
        set
        {
            if (value == null)
            {
                this.RemoveDateAxis();

                return;
            }

            if (value.Worksheet.Workbook != this._ws.Workbook)
            {
                throw new ArgumentException("Range must be in the same package");
            }
            else if (value.Rows != 1 && value.Columns != 1)
            {
                throw new ArgumentException("Range must only be 1 row or column");
            }

            this.DateAxis = true;
            this.SetXmlNodeString("xm:f", value.FullAddress);
        }
    }

    private void RemoveDateAxis()
    {
        this.DeleteNode("xm:f");
        this.DateAxis = false;
    }

    /// <summary>
    /// The range containing the data from the sparklines
    /// </summary>
    public ExcelRangeBase DataRange
    {
        get
        {
            if (this.Sparklines.Count == 0)
            {
                return null;
            }
            else
            {
                ExcelWorksheet? ws = this._ws.Workbook.Worksheets[this.Sparklines[0].RangeAddress.WorkSheetName];

                return this._ws.Cells[this.Sparklines[0].RangeAddress._fromRow,
                                      this.Sparklines[0].RangeAddress._fromCol,
                                      this.Sparklines[this.Sparklines.Count - 1].RangeAddress._toRow,
                                      this.Sparklines[this.Sparklines.Count - 1].RangeAddress._toCol];
            }
        }
    }

    /// <summary>
    /// The range containing the sparklines
    /// </summary>
    public ExcelRangeBase LocationRange
    {
        get
        {
            if (this.Sparklines.Count == 0)
            {
                return null;
            }
            else
            {
                return this._ws.Cells[this.Sparklines[0].Cell.Row,
                                      this.Sparklines[0].Cell.Column,
                                      this.Sparklines[this.Sparklines.Count - 1].Cell.Row,
                                      this.Sparklines[this.Sparklines.Count - 1].Cell.Column];
            }
        }
    }

    /// <summary>
    /// The Sparklines for the sparklinegroup
    /// </summary>
    public ExcelSparklineCollection Sparklines { get; internal set; }

    #region Boolean settings

    const string _dateAxisPath = "@dateAxis";

    internal bool DateAxis
    {
        get { return this.GetXmlNodeBool(_dateAxisPath, false); }
        set { this.SetXmlNodeBool(_dateAxisPath, value); }
    }

    const string _markersPath = "@markers";

    /// <summary>
    /// Highlight each point in each sparkline in the sparkline group.
    /// Applies to line sparklines only
    /// </summary>
    public bool Markers
    {
        get { return this.GetXmlNodeBool(_markersPath, false); }
        set { this.SetXmlNodeBool(_markersPath, value); }
    }

    const string _highPath = "@high";

    /// <summary>
    /// Highlight the highest point of data in the sparkline group
    /// </summary>
    public bool High
    {
        get { return this.GetXmlNodeBool(_highPath, false); }
        set { this.SetXmlNodeBool(_highPath, value); }
    }

    const string _lowPath = "@low";

    /// <summary>
    /// Highlight the lowest point of data in the sparkline group
    /// </summary>
    public bool Low
    {
        get { return this.GetXmlNodeBool(_lowPath, false); }
        set { this.SetXmlNodeBool(_lowPath, value); }
    }

    const string _firstPath = "@first";

    /// <summary>
    /// Highlight the first point of data in the sparkline group
    /// </summary>
    public bool First
    {
        get { return this.GetXmlNodeBool(_firstPath, false); }
        set { this.SetXmlNodeBool(_firstPath, value); }
    }

    const string _lastPath = "@last";

    /// <summary>
    /// Highlight the last point of data in the sparkline group
    /// </summary>
    public bool Last
    {
        get { return this.GetXmlNodeBool(_lastPath, false); }
        set { this.SetXmlNodeBool(_lastPath, value); }
    }

    const string _negativePath = "@negative";

    /// <summary>
    /// Highlight negative points of data in the sparkline group with a different color or marker
    /// </summary>
    public bool Negative
    {
        get { return this.GetXmlNodeBool(_negativePath); }
        set { this.SetXmlNodeBool(_negativePath, value); }
    }

    const string _displayXAxisPath = "@displayXAxis";

    /// <summary>
    /// Displays the X axis
    /// </summary>
    public bool DisplayXAxis
    {
        get { return this.GetXmlNodeBool(_displayXAxisPath); }
        set { this.SetXmlNodeBool(_displayXAxisPath, value); }
    }

    const string _displayHiddenPath = "@displayHidden";

    /// <summary>
    /// Display hidden cells
    /// </summary>
    public bool DisplayHidden
    {
        get { return this.GetXmlNodeBool(_displayHiddenPath); }
        set { this.SetXmlNodeBool(_displayHiddenPath, value); }
    }

    #endregion

    const string lineWidthPath = "@lineWeight";

    /// <summary>
    /// The weight of the line. Applies to line sparklines only.
    /// </summary>
    public double LineWidth
    {
        get { return this.GetXmlNodeDoubleNull(lineWidthPath) ?? 0.75; }
        set { this.SetXmlNodeString(lineWidthPath, value.ToString(CultureInfo.InvariantCulture)); }
    }

    const string _dispBlanksAsPath = "@displayEmptyCellsAs";

    /// <summary>
    /// How to display empty cells in the series
    /// </summary>
    public eDispBlanksAs DisplayEmptyCellsAs
    {
        get
        {
            string? s = this.GetXmlNodeString(_dispBlanksAsPath);

            if (string.IsNullOrEmpty(s))
            {
                return eDispBlanksAs.Zero;
            }
            else
            {
                return (eDispBlanksAs)Enum.Parse(typeof(eDispBlanksAs), s, true);
            }
        }
        set { this.SetXmlNodeString(_dispBlanksAsPath, value.ToString().ToLower()); }
    }

    const string _typePath = "@type";

    /// <summary>
    /// Type of sparkline
    /// </summary>
    public eSparklineType Type
    {
        get
        {
            string? type = this.GetXmlNodeString(_typePath);

            if (string.IsNullOrEmpty(type))
            {
                return eSparklineType.Line;
            }
            else
            {
                return (eSparklineType)Enum.Parse(typeof(eSparklineType), type, true);
            }
        }
        set { this.SetXmlNodeString(_typePath, value.ToString().ToLower()); }
    }

    #region Colors

    const string _colorSeriesPath = "x14:colorSeries";

    /// <summary>
    /// Sparkline color
    /// </summary>
    public ExcelSparklineColor ColorSeries
    {
        get
        {
            this.CreateNode(_colorSeriesPath);

            return new ExcelSparklineColor(this.NameSpaceManager, this.TopNode.SelectSingleNode(_colorSeriesPath, this.NameSpaceManager));
        }
    }

    const string _colorNegativePath = "x14:colorNegative";

    /// <summary>
    /// Markercolor for the lowest negative point
    /// </summary>  
    public ExcelSparklineColor ColorNegative
    {
        get
        {
            this.CreateNode(_colorNegativePath);

            return new ExcelSparklineColor(this.NameSpaceManager, this.TopNode.SelectSingleNode(_colorNegativePath, this.NameSpaceManager));
        }
    }

    const string _colorAxisPath = "x14:colorAxis";

    /// <summary>
    /// Markercolor for the lowest negative point
    /// </summary>
    public ExcelSparklineColor ColorAxis
    {
        get
        {
            this.CreateNode(_colorAxisPath);

            return new ExcelSparklineColor(this.NameSpaceManager, this.TopNode.SelectSingleNode(_colorAxisPath, this.NameSpaceManager));
        }
    }

    const string _colorMarkersPath = "x14:colorMarkers";

    /// <summary>
    /// Default marker color 
    /// </summary> 
    public ExcelSparklineColor ColorMarkers
    {
        get
        {
            this.CreateNode(_colorMarkersPath);

            return new ExcelSparklineColor(this.NameSpaceManager, this.TopNode.SelectSingleNode(_colorMarkersPath, this.NameSpaceManager));
        }
    }

    const string _colorFirstPath = "x14:colorFirst";

    /// <summary>
    /// The color of the first point
    /// </summary>
    public ExcelSparklineColor ColorFirst
    {
        get
        {
            this.CreateNode(_colorFirstPath);

            return new ExcelSparklineColor(this.NameSpaceManager, this.TopNode.SelectSingleNode(_colorFirstPath, this.NameSpaceManager));
        }
    }

    const string _colorLastPath = "x14:colorLast";

    /// <summary>
    /// The color of the last point
    /// </summary>
    public ExcelSparklineColor ColorLast
    {
        get
        {
            this.CreateNode(_colorLastPath);

            return new ExcelSparklineColor(this.NameSpaceManager, this.TopNode.SelectSingleNode(_colorLastPath, this.NameSpaceManager));
        }
    }

    const string _colorHighPath = "x14:colorHigh";

    /// <summary>
    /// The color of the point with the highest value
    /// </summary>
    public ExcelSparklineColor ColorHigh
    {
        get
        {
            this.CreateNode(_colorHighPath);

            return new ExcelSparklineColor(this.NameSpaceManager, this.TopNode.SelectSingleNode(_colorHighPath, this.NameSpaceManager));
        }
    }

    const string _colorLowPath = "x14:colorLow";

    /// <summary>
    /// The color of the point with the lowest value
    /// </summary>
    public ExcelSparklineColor ColorLow
    {
        get
        {
            this.CreateNode(_colorLowPath);

            return new ExcelSparklineColor(this.NameSpaceManager, this.TopNode.SelectSingleNode(_colorLowPath, this.NameSpaceManager));
        }
    }

    const string _manualMinPath = "@manualMin";

    /// <summary>
    /// When <c>MinAxisType</c> type is set to <c>Custom</c>, this value sets the minimum value
    /// <see cref="MinAxisType"/>
    /// </summary>
    public double ManualMin
    {
        get { return this.GetXmlNodeDouble(_manualMinPath); }
        set
        {
            this.SetXmlNodeString(_minAxisTypePath, "custom");
            this.SetXmlNodeString(_manualMinPath, value.ToString("F", CultureInfo.InvariantCulture));
        }
    }

    const string _manualMaxPath = "@manualMax";

    /// <summary>
    /// When <c>MaxAxisType</c> type is set to <c>Custom</c>, this value sets the maximum value
    /// <see cref="MaxAxisType"/>
    /// </summary>
    public double ManualMax
    {
        get { return this.GetXmlNodeDouble(_manualMaxPath); }
        set
        {
            this.SetXmlNodeString(_maxAxisTypePath, "custom");
            this.SetXmlNodeString(_manualMaxPath, value.ToString("F", CultureInfo.InvariantCulture));
        }
    }

    const string _minAxisTypePath = "@minAxisType";

    /// <summary>
    /// Vertical axis minimum value options
    /// </summary>
    public eSparklineAxisMinMax MinAxisType
    {
        get
        {
            string? s = this.GetXmlNodeString(_minAxisTypePath);

            if (string.IsNullOrEmpty(s))
            {
                return eSparklineAxisMinMax.Individual;
            }
            else
            {
                return (eSparklineAxisMinMax)Enum.Parse(typeof(eSparklineAxisMinMax), s, true);
            }
        }
        set
        {
            if (value == eSparklineAxisMinMax.Custom)
            {
                this.ManualMin = 0;
            }
            else
            {
                this.SetXmlNodeString(_minAxisTypePath, value.ToEnumString());
                this.DeleteNode(_manualMinPath);
            }
        }
    }

    const string _maxAxisTypePath = "@maxAxisType";

    /// <summary>
    /// Vertical axis maximum value options
    /// </summary>
    public eSparklineAxisMinMax MaxAxisType
    {
        get
        {
            string? s = this.GetXmlNodeString(_maxAxisTypePath);

            if (string.IsNullOrEmpty(s))
            {
                return eSparklineAxisMinMax.Individual;
            }
            else
            {
                return (eSparklineAxisMinMax)Enum.Parse(typeof(eSparklineAxisMinMax), s, true);
            }
        }
        set
        {
            if (value == eSparklineAxisMinMax.Custom)
            {
                this.ManualMax = 0;
            }
            else
            {
                this.SetXmlNodeString(_maxAxisTypePath, value.ToEnumString());
                this.DeleteNode(_manualMaxPath);
            }
        }
    }

    const string _rightToLeftPath = "@rightToLeft";

    /// <summary>
    /// Plot horizontal axis data right to left if true
    /// </summary>
    public bool RightToLeft
    {
        get { return this.GetXmlNodeBool(_rightToLeftPath, false); }
        set { this.SetXmlNodeBool(_rightToLeftPath, value); }
    }

    #endregion
}