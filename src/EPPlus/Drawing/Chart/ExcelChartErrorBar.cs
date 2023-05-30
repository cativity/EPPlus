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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// The title of a chart
/// </summary>
public class ExcelChartErrorBars : XmlHelper, IDrawingStyleBase
{
    readonly ExcelChartSerieWithErrorBars _chartSerie;

    internal ExcelChartErrorBars(ExcelChartSerieWithErrorBars chartSerie)
        : this(chartSerie, chartSerie.TopNode)
    {
    }

    internal ExcelChartErrorBars(ExcelChartSerieWithErrorBars chartSerie, XmlNode topNode)
        : base(chartSerie.NameSpaceManager, topNode)
    {
        this._chartSerie = chartSerie;

        this.AddSchemaNodeOrder(new string[] { "errDir", "errBarType", "errValType", "noEndCap", "plus", "minus", "val", "spPr" },
                                ExcelDrawing._schemaNodeOrderSpPr);

        if (this.TopNode.LocalName != "errBars")
        {
            this.TopNode = chartSerie.CreateNode("c:errBars", false, true);
        }
    }

    string _directionPath = "c:errDir/@val";

    /// <summary>
    /// The directions for the error bars. For scatter-, bubble- and area charts this property can't be changed. Please use the ErrorBars property for Y direction and ErrorBarsX for the X direction.
    /// </summary>
    public eErrorBarDirection Direction
    {
        get
        {
            this.ValidateNotDeleted();

            return this.GetXmlNodeString(this._directionPath).ToEnum(eErrorBarDirection.Y);
        }
        set
        {
            this.ValidateNotDeleted();

            if (this._chartSerie._chart.IsTypeBubble() || this._chartSerie._chart.IsTypeScatter() || this._chartSerie._chart.IsTypeArea())
            {
                if (value != this.Direction)
                {
                    throw new
                        InvalidOperationException("Can't change direction for this chart type. Please use ErrorBars or ErrorBarsX property to determine the direction.");
                }

                return;
            }

            this.SetDirection(value);
        }
    }

    internal void SetDirection(eErrorBarDirection value)
    {
        this.SetXmlNodeString(this._directionPath, value.ToEnumString());
    }

    string _barTypePath = "c:errBarType/@val";

    /// <summary>
    /// The ways to draw an error bar
    /// </summary>
    public eErrorBarType BarType
    {
        get
        {
            this.ValidateNotDeleted();

            return this.GetXmlNodeString(this._barTypePath).ToEnum(eErrorBarType.Both);
        }
        set
        {
            this.ValidateNotDeleted();
            this.SetXmlNodeString(this._barTypePath, value.ToEnumString());
        }
    }

    string _valueTypePath = "c:errValType/@val";

    /// <summary>
    /// The ways to determine the length of the error bars
    /// </summary>
    public eErrorValueType ValueType
    {
        get
        {
            this.ValidateNotDeleted();

            return this.GetXmlNodeString(this._valueTypePath).TranslateErrorValueType();
        }
        set
        {
            this.ValidateNotDeleted();
            this.SetXmlNodeString(this._valueTypePath, value.ToEnumString());
        }
    }

    string _noEndCapPath = "c:noEndCap/@val";

    /// <summary>
    /// If true, no end cap is drawn on the error bars 
    /// </summary>
    public bool NoEndCap
    {
        get
        {
            this.ValidateNotDeleted();

            return this.GetXmlNodeBool(this._noEndCapPath, true);
        }
        set
        {
            this.ValidateNotDeleted();
            this.SetXmlNodeBool(this._noEndCapPath, value, true);
        }
    }

    string _valuePath = "c:val/@val";

    /// <summary>
    /// The value which used to determine the length of the error bars when <c>ValueType</c> is FixedValue
    /// </summary>
    public double? Value
    {
        get
        {
            this.ValidateNotDeleted();

            return this.GetXmlNodeDoubleNull(this._valuePath);
        }
        set
        {
            this.ValidateNotDeleted();

            if (value == null)
            {
                this.DeleteNode(this._valuePath, true);
            }
            else
            {
                this.SetXmlNodeString(this._valuePath, value.Value.ToString("R15", CultureInfo.InvariantCulture));
            }
        }
    }

    string _plusNodePath = "c:plus";
    ExcelChartNumericSource _plus;

    /// <summary>
    /// Numeric Source for plus errorbars when <c>ValueType</c> is set to Custom
    /// </summary>
    public ExcelChartNumericSource Plus
    {
        get
        {
            this.ValidateNotDeleted();

            return this._plus ??= new ExcelChartNumericSource(this.NameSpaceManager, this.TopNode, this._plusNodePath, this.SchemaNodeOrder);
        }
    }

    string _minusNodePath = "c:minus";
    ExcelChartNumericSource _minus;

    /// <summary>
    /// Numeric Source for minus errorbars when <c>ValueType</c> is set to Custom
    /// </summary>
    public ExcelChartNumericSource Minus
    {
        get
        {
            this.ValidateNotDeleted();

            return this._minus ??= new ExcelChartNumericSource(this.NameSpaceManager, this.TopNode, this._minusNodePath, this.SchemaNodeOrder);
        }
    }

    ExcelDrawingFill _fill;

    /// <summary>
    /// Fill style
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            this.ValidateNotDeleted();

            return this._fill ??= new ExcelDrawingFill(this._chartSerie._chart, this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);
        }
    }

    ExcelDrawingBorder _border;

    /// <summary>
    /// Border style
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get
        {
            this.ValidateNotDeleted();

            return this._border ??= new ExcelDrawingBorder(this._chartSerie._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:ln", this.SchemaNodeOrder);
        }
    }

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            this.ValidateNotDeleted();

            return this._effect ??= new ExcelDrawingEffectStyle(this._chartSerie._chart,
                                                                this.NameSpaceManager,
                                                                this.TopNode,
                                                                "c:spPr/a:effectLst",
                                                                this.SchemaNodeOrder);
        }
    }

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get
        {
            this.ValidateNotDeleted();

            return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);
        }
    }

    private void ValidateNotDeleted()
    {
        if (this.TopNode == null)
        {
            throw new InvalidOperationException("The error bar has been deleted.");
        }
    }

    void IDrawingStyleBase.CreatespPr()
    {
        this.CreatespPrNode();
    }

    /// <summary>
    /// Remove the error bars
    /// </summary>
    public void Remove()
    {
        this.DeleteNode(".");

        if (this._chartSerie.ErrorBars == this)
        {
            this._chartSerie.ErrorBars = null;
        }

        if (this._chartSerie is ExcelChartSerieWithHorizontalErrorBars errorBarsSerie)
        {
            if (errorBarsSerie.ErrorBarsX == this)
            {
                errorBarsSerie.ErrorBarsX = null;
            }
        }
    }
}