/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/

using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Controls;

/// <summary>
/// Margin setting for a vml drawing
/// </summary>
public class ExcelControlMargin
{
    private ExcelControlWithText _control;
    private XmlHelper _vmlHelper;

    internal ExcelControlMargin(ExcelControlWithText control)
    {
        this._control = control;
        this._vmlHelper = XmlHelperFactory.Create(control._vmlProp.NameSpaceManager, control._vmlProp.TopNode.ParentNode);

        this.Automatic = this._vmlHelper.GetXmlNodeString("@o:insetmode") == "auto";
        string? margin = this._vmlHelper.GetXmlNodeString("v:textbox/@inset");

        string? v = margin.GetCsvPosition(0);
        this.LeftMargin.SetValue(v);

        v = margin.GetCsvPosition(1);
        this.TopMargin.SetValue(v);

        v = margin.GetCsvPosition(2);
        this.RightMargin.SetValue(v);

        v = margin.GetCsvPosition(3);
        this.BottomMargin.SetValue(v);
    }

    /// <summary>
    /// Sets the margin value and unit of measurement for all margins.
    /// </summary>
    /// <param name="marginValue">Margin value to set for all margins</param>
    /// <param name="unit">The unit to set for all margins. Default <see cref="eMeasurementUnits.Points" /></param>
    public void SetValue(double marginValue, eMeasurementUnits unit = eMeasurementUnits.Points)
    {
        this.LeftMargin.Value = marginValue;
        this.TopMargin.Value = marginValue;
        this.RightMargin.Value = marginValue;
        this.BottomMargin.Value = marginValue;
        this.SetUnit(unit);
    }

    /// <summary>
    /// Sets the margin unit of measurement for all margins.
    /// </summary>
    /// <param name="unit">The unit to set for all margins.</param>
    public void SetUnit(eMeasurementUnits unit)
    {
        this.LeftMargin.Unit = unit;
        this.TopMargin.Unit = unit;
        this.RightMargin.Unit = unit;
        this.BottomMargin.Unit = unit;
    }

    internal void UpdateXml()
    {
        if (this.Automatic)
        {
            this._vmlHelper.SetXmlNodeString("@o:insetmode", "auto");
        }
        else
        {
            this._vmlHelper.DeleteNode("@o:insetmode"); //Custom
        }

        if (this.LeftMargin.Value != 0 && this.TopMargin.Value != 0 && this.RightMargin.Value != 0 && this.BottomMargin.Value != 0)
        {
            string? v = this.LeftMargin.GetValueString()
                        + ","
                        + this.TopMargin.GetValueString()
                        + ","
                        + this.RightMargin.GetValueString()
                        + ","
                        + this.BottomMargin.GetValueString();

            this._control.TextBody.LeftInsert = this.LeftMargin.ToEmu();
            this._control.TextBody.TopInsert = this.TopMargin.ToEmu();
            this._control.TextBody.RightInsert = this.RightMargin.ToEmu();
            this._control.TextBody.BottomInsert = this.BottomMargin.ToEmu();

            this._vmlHelper.SetXmlNodeString("v:textbox/@inset", v);
        }
        else
        {
            this._vmlHelper.DeleteNode("v:textbox/@inset");
        }
    }

    /// <summary>
    /// Margin is autiomatic
    /// </summary>
    public bool Automatic { get; set; }

    /// <summary>
    /// Left Margin
    /// </summary>
    public ExcelVmlMeasurementUnit LeftMargin { get; } = new ExcelVmlMeasurementUnit();

    /// <summary>
    /// Right Margin
    /// </summary>
    public ExcelVmlMeasurementUnit RightMargin { get; } = new ExcelVmlMeasurementUnit();

    /// <summary>
    /// Top Margin
    /// </summary>
    public ExcelVmlMeasurementUnit TopMargin { get; } = new ExcelVmlMeasurementUnit();

    /// <summary>
    /// Bottom margin
    /// </summary>
    public ExcelVmlMeasurementUnit BottomMargin { get; } = new ExcelVmlMeasurementUnit();
}