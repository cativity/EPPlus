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
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.ThreeD;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A trendline object
/// </summary>
public class ExcelChartTrendline : XmlHelper, IDrawingStyleBase
{
    ExcelChartStandardSerie _serie;

    internal ExcelChartTrendline(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelChartStandardSerie serie)
        : base(namespaceManager, topNode)

    {
        this._serie = serie;

        this.AddSchemaNodeOrder(new string[]
                                {
                                    "name", "spPr", "trendlineType", "order", "period", "forward", "backward", "intercept", "dispRSqr", "dispEq",
                                    "trendlineLbl"
                                },
                                ExcelDrawing._schemaNodeOrderSpPr);
    }

    const string TRENDLINEPATH = "c:trendlineType/@val";

    /// <summary>
    /// Type of Trendline
    /// </summary>
    public eTrendLine Type
    {
        get
        {
            switch (this.GetXmlNodeString(TRENDLINEPATH).ToLower(CultureInfo.InvariantCulture))
            {
                case "exp":
                    return eTrendLine.Exponential;

                case "log":
                    return eTrendLine.Logarithmic;

                case "poly":
                    return eTrendLine.Polynomial;

                case "movingavg":
                    return eTrendLine.MovingAvgerage;

                case "power":
                    return eTrendLine.Power;

                default:
                    return eTrendLine.Linear;
            }
        }
        set
        {
            switch (value)
            {
                case eTrendLine.Exponential:
                    this.SetXmlNodeString(TRENDLINEPATH, "exp");

                    break;

                case eTrendLine.Logarithmic:
                    this.SetXmlNodeString(TRENDLINEPATH, "log");

                    break;

                case eTrendLine.Polynomial:
                    this.SetXmlNodeString(TRENDLINEPATH, "poly");
                    this.Order = 2;

                    break;

                case eTrendLine.MovingAvgerage:
                    this.SetXmlNodeString(TRENDLINEPATH, "movingAvg");
                    this.Period = 2;

                    break;

                case eTrendLine.Power:
                    this.SetXmlNodeString(TRENDLINEPATH, "power");

                    break;

                default:
                    this.SetXmlNodeString(TRENDLINEPATH, "linear");

                    break;
            }
        }
    }

    const string NAMEPATH = "c:name";

    /// <summary>
    /// Name in the legend
    /// </summary>
    public string Name
    {
        get { return this.GetXmlNodeString(NAMEPATH); }
        set { this.SetXmlNodeString(NAMEPATH, value, true); }
    }

    const string ORDERPATH = "c:order/@val";

    /// <summary>
    /// Order for polynominal trendlines
    /// </summary>
    public decimal Order
    {
        get { return this.GetXmlNodeDecimal(ORDERPATH); }
        set
        {
            if (this.Type == eTrendLine.MovingAvgerage)
            {
                throw new ArgumentException("Can't set period for trendline type MovingAvgerage");
            }

            this.DeleteAllNode(PERIODPATH);
            this.SetXmlNodeString(ORDERPATH, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string PERIODPATH = "c:period/@val";

    /// <summary>
    /// Period for monthly average trendlines
    /// </summary>
    public decimal Period
    {
        get { return this.GetXmlNodeDecimal(PERIODPATH); }
        set
        {
            if (this.Type == eTrendLine.Polynomial)
            {
                throw new ArgumentException("Can't set period for trendline type Polynomial");
            }

            this.DeleteAllNode(ORDERPATH);
            this.SetXmlNodeString(PERIODPATH, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string FORWARDPATH = "c:forward/@val";

    /// <summary>
    /// Forcast forward periods
    /// </summary>
    public decimal Forward
    {
        get { return this.GetXmlNodeDecimal(FORWARDPATH); }
        set { this.SetXmlNodeString(FORWARDPATH, value.ToString(CultureInfo.InvariantCulture)); }
    }

    const string BACKWARDPATH = "c:backward/@val";

    /// <summary>
    /// Forcast backwards periods
    /// </summary>
    public decimal Backward
    {
        get { return this.GetXmlNodeDecimal(BACKWARDPATH); }
        set { this.SetXmlNodeString(BACKWARDPATH, value.ToString(CultureInfo.InvariantCulture)); }
    }

    const string INTERCEPTPATH = "c:intercept/@val";

    /// <summary>
    /// The point where the trendline crosses the vertical axis
    /// </summary>
    public decimal Intercept
    {
        get { return this.GetXmlNodeDecimal(INTERCEPTPATH); }
        set { this.SetXmlNodeString(INTERCEPTPATH, value.ToString(CultureInfo.InvariantCulture)); }
    }

    const string DISPLAYRSQUAREDVALUEPATH = "c:dispRSqr/@val";

    /// <summary>
    /// If to display the R-squared value for a trendline
    /// </summary>
    public bool DisplayRSquaredValue
    {
        get { return this.GetXmlNodeBool(DISPLAYRSQUAREDVALUEPATH, true); }
        set { this.SetXmlNodeBool(DISPLAYRSQUAREDVALUEPATH, value, true); }
    }

    const string DISPLAYEQUATIONPATH = "c:dispEq/@val";

    /// <summary>
    /// If to display the trendline equation on the chart
    /// </summary>
    public bool DisplayEquation
    {
        get { return this.GetXmlNodeBool(DISPLAYEQUATIONPATH, true); }
        set { this.SetXmlNodeBool(DISPLAYEQUATIONPATH, value, true); }
    }

    ExcelDrawingFill _fill = null;

    /// <summary>
    /// Access to fill properties
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get { return this._fill ??= new ExcelDrawingFill(this._serie._chart, this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder); }
    }

    ExcelDrawingBorder _border = null;

    /// <summary>
    /// Access to border properties
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get { return this._border ??= new ExcelDrawingBorder(this._serie._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:ln", this.SchemaNodeOrder); }
    }

    ExcelDrawingEffectStyle _effect = null;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            return this._effect ??= new ExcelDrawingEffectStyle(this._serie._chart,
                                                                this.NameSpaceManager,
                                                                this.TopNode,
                                                                "c:spPr/a:effectLst",
                                                                this.SchemaNodeOrder);
        }
    }

    ExcelDrawing3D _threeD = null;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get { return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder); }
    }

    void IDrawingStyleBase.CreatespPr()
    {
        this.CreatespPrNode();
    }

    ExcelChartTrendlineLabel _label = null;

    /// <summary>
    /// Trendline labels
    /// </summary>
    public ExcelChartTrendlineLabel Label
    {
        get { return this._label ??= new ExcelChartTrendlineLabel(this.NameSpaceManager, this.TopNode, this._serie); }
    }

    /// <summary>
    /// Return true if the trendline has labels.
    /// </summary>
    public bool HasLbl
    {
        get
        {
            return this.ExistsNode("c:trendlineLbl")
                   || (this.Type != eTrendLine.MovingAvgerage && (this.DisplayRSquaredValue == true || this.DisplayEquation == true));
        }
    }
}