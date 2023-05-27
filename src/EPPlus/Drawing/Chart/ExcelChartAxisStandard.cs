/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/22/2020         EPPlus Software AB       Added this class
 *************************************************************************************************/

using System;
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// An axis for a standard chart. 
/// </summary>
public sealed class ExcelChartAxisStandard : ExcelChartAxis
{
    internal readonly static string[] _schemaNodeOrderDateShared = new string[]
    {
        "axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr",
        "txPr", "crossAx", "crosses", "crossesAt"
    };

    internal static string[] _schemaNodeOrderCat = new string[] { "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "extLst" };

    internal static string[] _schemaNodeOrderDate =
        new string[] { "auto", "lblOffset", "baseTimeUnit", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit", "extLst" };

    internal static string[] _schemaNodeOrderSer = new string[] { "tickLblSkip", "tickMarkSkip", "extLst" };
    internal static string[] _schemaNodeOrderVal = new string[] { "crossBetween", "majorUnit", "minorUnit", "dispUnits", "extLst" };

    internal ExcelChartAxisStandard(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string nsPrefix)
        : base(chart, nameSpaceManager, topNode, nsPrefix)
    {
        this.AddSchemaNodeOrder(new string[]
                                {
                                    "axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark",
                                    "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "crossBetween", "auto", "lblOffset",
                                    "baseTimeUnit", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit", "tickLblSkip", "tickMarkSkip", "dispUnits",
                                    "noMultiLvlLbl", "logBase", "orientation", "max", "min"
                                },
                                ExcelDrawing._schemaNodeOrderSpPr);
    }

    internal override string Id
    {
        get { return this.GetXmlNodeString("c:axId/@val"); }
    }

    const string _majorTickMark = "c:majorTickMark/@val";

    /// <summary>
    /// Get or Sets the major tick marks for the axis. 
    /// </summary>
    public override eAxisTickMark MajorTickMark
    {
        get
        {
            string? v = this.GetXmlNodeString(_majorTickMark);

            if (string.IsNullOrEmpty(v))
            {
                return eAxisTickMark.Cross;
            }
            else
            {
                try
                {
                    return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), v);
                }
                catch
                {
                    return eAxisTickMark.Cross;
                }
            }
        }
        set { this.SetXmlNodeString(_majorTickMark, value.ToString().ToLower(CultureInfo.InvariantCulture)); }
    }

    const string _minorTickMark = "c:minorTickMark/@val";

    /// <summary>
    /// Get or Sets the minor tick marks for the axis. 
    /// </summary>
    public override eAxisTickMark MinorTickMark
    {
        get
        {
            string? v = this.GetXmlNodeString(_minorTickMark);

            if (string.IsNullOrEmpty(v))
            {
                return eAxisTickMark.Cross;
            }
            else
            {
                try
                {
                    return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), v);
                }
                catch
                {
                    return eAxisTickMark.Cross;
                }
            }
        }
        set { this.SetXmlNodeString(_minorTickMark, value.ToString().ToLower(CultureInfo.InvariantCulture)); }
    }

    private string AXIS_POSITION_PATH = "c:axPos/@val";

    /// <summary>
    /// Where the axis is located
    /// </summary>
    public override eAxisPosition AxisPosition
    {
        get
        {
            switch (this.GetXmlNodeString(this.AXIS_POSITION_PATH))
            {
                case "b":
                    return eAxisPosition.Bottom;

                case "r":
                    return eAxisPosition.Right;

                case "t":
                    return eAxisPosition.Top;

                default:
                    return eAxisPosition.Left;
            }
        }
        internal set { this.SetXmlNodeString(this.AXIS_POSITION_PATH, value.ToString().ToLower(CultureInfo.InvariantCulture).Substring(0, 1)); }
    }

    /// <summary>
    /// Chart axis title
    /// </summary>
    public new ExcelChartTitleStandard Title
    {
        get { return (ExcelChartTitleStandard)this.GetTitle(); }
    }

    internal override ExcelChartTitle GetTitle()
    {
        if (this._title == null)
        {
            this.AddTitleNode();
            this._title = new ExcelChartTitleStandard(this._chart, this.NameSpaceManager, this.TopNode, "c");
        }

        return this._title;
    }

    const string _minValuePath = "c:scaling/c:min/@val";

    /// <summary>
    /// Minimum value for the axis.
    /// Null is automatic
    /// </summary>
    public override double? MinValue
    {
        get { return this.GetXmlNodeDoubleNull(_minValuePath); }
        set
        {
            if (value == null)
            {
                this.DeleteNode(_minValuePath, true);
            }
            else
            {
                this.SetXmlNodeString(_minValuePath, ((double)value).ToString(CultureInfo.InvariantCulture));
            }
        }
    }

    const string _maxValuePath = "c:scaling/c:max/@val";

    /// <summary>
    /// Max value for the axis.
    /// Null is automatic
    /// </summary>
    public override double? MaxValue
    {
        get { return this.GetXmlNodeDoubleNull(_maxValuePath); }
        set
        {
            if (value == null)
            {
                this.DeleteNode(_maxValuePath, true);
            }
            else
            {
                this.SetXmlNodeString(_maxValuePath, ((double)value).ToString(CultureInfo.InvariantCulture));
            }
        }
    }

    const string _lblPos = "c:tickLblPos/@val";

    /// <summary>
    /// The Position of the labels
    /// </summary>
    public override eTickLabelPosition LabelPosition
    {
        get
        {
            string? v = this.GetXmlNodeString(_lblPos);

            if (string.IsNullOrEmpty(v))
            {
                return eTickLabelPosition.NextTo;
            }
            else
            {
                try
                {
                    return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), v, true);
                }
                catch
                {
                    return eTickLabelPosition.NextTo;
                }
            }
        }
        set
        {
            string lp = value.ToString();
            this.SetXmlNodeString(_lblPos, lp.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + lp.Substring(1, lp.Length - 1));
        }
    }

    const string _crossesPath = "c:crosses/@val";

    /// <summary>
    /// Where the axis crosses
    /// </summary>
    public override eCrosses Crosses
    {
        get
        {
            string? v = this.GetXmlNodeString(_crossesPath);

            if (string.IsNullOrEmpty(v))
            {
                return eCrosses.AutoZero;
            }
            else
            {
                try
                {
                    return (eCrosses)Enum.Parse(typeof(eCrosses), v, true);
                }
                catch
                {
                    return eCrosses.AutoZero;
                }
            }
        }
        set
        {
            string? v = value.ToString();
            v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
            this.SetXmlNodeString(_crossesPath, v);
        }
    }

    const string _crossBetweenPath = "c:crossBetween/@val";

    /// <summary>
    /// How the axis are crossed
    /// </summary>
    public override eCrossBetween CrossBetween
    {
        get
        {
            string? v = this.GetXmlNodeString(_crossBetweenPath);

            if (string.IsNullOrEmpty(v))
            {
                return eCrossBetween.Between;
            }
            else
            {
                try
                {
                    return (eCrossBetween)Enum.Parse(typeof(eCrossBetween), v, true);
                }
                catch
                {
                    return eCrossBetween.Between;
                }
            }
        }
        set
        {
            string? v = value.ToString();
            v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1);
            this.SetXmlNodeString(_crossBetweenPath, v);
        }
    }

    const string _crossesAtPath = "c:crossesAt/@val";

    /// <summary>
    /// The value where the axis cross. 
    /// Null is automatic
    /// </summary>
    public override double? CrossesAt
    {
        get { return this.GetXmlNodeDoubleNull(_crossesAtPath); }
        set
        {
            if (value == null)
            {
                this.DeleteNode(_crossesAtPath, true);
            }
            else
            {
                this.SetXmlNodeString(_crossesAtPath, ((double)value).ToString(CultureInfo.InvariantCulture));
            }
        }
    }

    /// <summary>
    /// If the axis is deleted
    /// </summary>
    public override bool Deleted
    {
        get { return this.GetXmlNodeBool("c:delete/@val"); }
        set { this.SetXmlNodeBool("c:delete/@val", value); }
    }

    const string _ticLblPos_Path = "c:tickLblPos/@val";

    /// <summary>
    /// Position of the Lables
    /// </summary>
    public override eTickLabelPosition TickLabelPosition
    {
        get
        {
            string v = this.GetXmlNodeString(_ticLblPos_Path);

            if (v == "")
            {
                return eTickLabelPosition.None;
            }
            else
            {
                return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), v, true);
            }
        }
        set
        {
            string v = value.ToString();
            v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
            this.SetXmlNodeString(_ticLblPos_Path, v);
        }
    }

    const string _displayUnitPath = "c:dispUnits/c:builtInUnit/@val";
    const string _custUnitPath = "c:dispUnits/c:custUnit/@val";

    /// <summary>
    /// The scaling value of the display units for the value axis
    /// </summary>
    public override double DisplayUnit
    {
        get
        {
            string v = this.GetXmlNodeString(_displayUnitPath);

            if (string.IsNullOrEmpty(v))
            {
                double? c = this.GetXmlNodeDoubleNull(_custUnitPath);

                if (c == null)
                {
                    return 0;
                }
                else
                {
                    return c.Value;
                }
            }
            else
            {
                try
                {
                    return (double)(long)Enum.Parse(typeof(eBuildInUnits), v, true);
                }
                catch
                {
                    return 0;
                }
            }
        }
        set
        {
            if (this.AxisType == eAxisType.Val && value >= 0)
            {
                foreach (object? v in Enum.GetValues(typeof(eBuildInUnits)))
                {
                    if ((double)(long)v == value)
                    {
                        this.DeleteNode(_custUnitPath, true);
                        this.SetXmlNodeString(_displayUnitPath, ((eBuildInUnits)value).ToString());

                        return;
                    }
                }

                this.DeleteNode(_displayUnitPath, true);

                if (value != 0)
                {
                    this.SetXmlNodeString(_custUnitPath, value.ToString(CultureInfo.InvariantCulture));
                }
            }
        }
    }

    const string _majorUnitPath = "c:majorUnit/@val";
    const string _majorUnitCatPath = "c:tickLblSkip/@val";

    /// <summary>
    /// Major unit for the axis.
    /// Null is automatic
    /// </summary>
    public override double? MajorUnit
    {
        get
        {
            if (this.AxisType == eAxisType.Cat)
            {
                return this.GetXmlNodeDoubleNull(_majorUnitCatPath);
            }
            else
            {
                return this.GetXmlNodeDoubleNull(_majorUnitPath);
            }
        }
        set
        {
            if (value == null)
            {
                this.DeleteNode(_majorUnitPath, true);
                this.DeleteNode(_majorUnitCatPath, true);
            }
            else
            {
                if (this.AxisType == eAxisType.Cat)
                {
                    this.SetXmlNodeString(_majorUnitCatPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    this.SetXmlNodeString(_majorUnitPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
    }

    const string _majorTimeUnitPath = "c:majorTimeUnit/@val";

    /// <summary>
    /// Major time unit for the axis.
    /// Null is automatic
    /// </summary>
    public override eTimeUnit? MajorTimeUnit
    {
        get
        {
            string? v = this.GetXmlNodeString(_majorTimeUnitPath);

            if (string.IsNullOrEmpty(v))
            {
                return null;
            }
            else
            {
                return v.ToEnum(eTimeUnit.Years);
            }
        }
        set
        {
            if (value.HasValue)
            {
                this.SetXmlNodeString(_majorTimeUnitPath, value.ToEnumString());
            }
            else
            {
                this.DeleteNode(_majorTimeUnitPath, true);
            }
        }
    }

    const string _minorUnitPath = "c:minorUnit/@val";
    const string _minorUnitCatPath = "c:tickMarkSkip/@val";

    /// <summary>
    /// Minor unit for the axis.
    /// Null is automatic
    /// </summary>
    public override double? MinorUnit
    {
        get
        {
            if (this.AxisType == eAxisType.Cat)
            {
                return this.GetXmlNodeDoubleNull(_minorUnitCatPath);
            }
            else
            {
                return this.GetXmlNodeDoubleNull(_minorUnitPath);
            }
        }
        set
        {
            if (value == null)
            {
                this.DeleteNode(_minorUnitPath, true);
                this.DeleteNode(_minorUnitCatPath, true);
            }
            else
            {
                if (this.AxisType == eAxisType.Cat)
                {
                    this.SetXmlNodeString(_minorUnitCatPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    this.SetXmlNodeString(_minorUnitPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
    }

    const string _minorTimeUnitPath = "c:minorTimeUnit/@val";

    /// <summary>
    /// Minor time unit for the axis.
    /// Null is automatic
    /// </summary>
    public override eTimeUnit? MinorTimeUnit
    {
        get
        {
            string? v = this.GetXmlNodeString(_minorTimeUnitPath);

            if (string.IsNullOrEmpty(v))
            {
                return null;
            }
            else
            {
                return v.ToEnum(eTimeUnit.Years);
            }
        }
        set
        {
            if (value.HasValue)
            {
                this.SetXmlNodeString(_minorTimeUnitPath, value.ToEnumString());
            }
            else
            {
                this.DeleteNode(_minorTimeUnitPath, true);
            }
        }
    }

    const string _logbasePath = "c:scaling/c:logBase/@val";

    /// <summary>
    /// The base for a logaritmic scale
    /// Null for a normal scale
    /// </summary>
    public override double? LogBase
    {
        get { return this.GetXmlNodeDoubleNull(_logbasePath); }
        set
        {
            if (value == null)
            {
                this.DeleteNode(_logbasePath, true);
            }
            else
            {
                double v = (double)value;

                if (v < 2 || v > 1000)
                {
                    throw new ArgumentOutOfRangeException("Value must be between 2 and 1000");
                }

                this.SetXmlNodeString(_logbasePath, v.ToString("0.0", CultureInfo.InvariantCulture));
            }
        }
    }

    const string _orientationPath = "c:scaling/c:orientation/@val";

    /// <summary>
    /// Axis orientation
    /// </summary>
    public override eAxisOrientation Orientation
    {
        get
        {
            string v = this.GetXmlNodeString(_orientationPath);

            if (v == "")
            {
                return eAxisOrientation.MinMax;
            }
            else
            {
                return (eAxisOrientation)Enum.Parse(typeof(eAxisOrientation), v, true);
            }
        }
        set
        {
            string s = value.ToString();
            s = s.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1, s.Length - 1);
            this.SetXmlNodeString(_orientationPath, s);
        }
    }

    internal override eAxisType AxisType
    {
        get
        {
            try
            {
                string? axType = this.TopNode.LocalName.Substring(0, this.TopNode.LocalName.Length - 2);

                if (axType == "ser")
                {
                    return eAxisType.Serie;
                }

                return (eAxisType)Enum.Parse(typeof(eAxisType), axType, true);
            }
            catch
            {
                return eAxisType.Val;
            }
        }
    }

    /// <summary>
    /// Adds the axis title and styles it according to the style selected in the StyleManager
    /// </summary>
    /// <param name="title"></param>
    public void AddTitle(ExcelRangeBase linkedCell)
    {
        this.Title.LinkedCell = linkedCell;
        this._chart.ApplyStyleOnPart(this.Title, this._chart._styleManager?.Style?.AxisTitle);
    }
}