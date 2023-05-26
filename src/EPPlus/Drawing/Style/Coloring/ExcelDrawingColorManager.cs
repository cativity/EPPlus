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
using System.Xml;
using System;
using System.Linq;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.Style.Coloring;

/// <summary>
/// Handles colors for drawings
/// </summary>
public class ExcelDrawingColorManager : ExcelDrawingThemeColorManager
{
    internal ExcelDrawingColorManager(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder, Action initMethod = null) : 
        base(nameSpaceManager, topNode, path, schemaNodeOrder, initMethod)
    {
        if (this._pathNode == null || this._colorNode==null)
        {
            return;
        }

        switch (this._colorNode.LocalName)
        {
            case "schemeClr":
                this.ColorType = eDrawingColorType.Scheme;
                this.SchemeColor = new ExcelDrawingSchemeColor(this._nameSpaceManager, this._colorNode);
                break;
        }
    }
    /// <summary>
    /// If <c>type</c> is set to SchemeColor, then this property contains the scheme color
    /// </summary>
    public ExcelDrawingSchemeColor SchemeColor { get; private set; }
    /// <summary>
    /// Set the color to a scheme color
    /// </summary>
    /// <param name="schemeColor">The scheme color</param>
    public void SetSchemeColor(eSchemeColor schemeColor)
    {
        this.ColorType = eDrawingColorType.Scheme;
        this.ResetColors(ExcelDrawingSchemeColor.NodeName);
        this.SchemeColor = new ExcelDrawingSchemeColor(this._nameSpaceManager, this._colorNode) { Color=schemeColor };
    }
    /// <summary>
    /// Reset the colors on the object
    /// </summary>
    /// <param name="newNodeName">The new color new name</param>
    internal new protected void ResetColors(string newNodeName) 
    {
        base.ResetColors(newNodeName);
        this.SchemeColor = null;
    }

    internal void ApplyNewColor(ExcelDrawingColorManager newColor, ExcelColorTransformCollection variation=null)
    {
        this.ColorType = newColor.ColorType;
        switch (newColor.ColorType)
        {
            case eDrawingColorType.Rgb:
                this.SetRgbColor(newColor.RgbColor.Color);
                break;
            case eDrawingColorType.RgbPercentage:
                this.SetRgbPercentageColor(newColor.RgbPercentageColor.RedPercentage, newColor.RgbPercentageColor.GreenPercentage, newColor.RgbPercentageColor.BluePercentage);
                break;
            case eDrawingColorType.Hsl:
                this.SetHslColor(newColor.HslColor.Hue, newColor.HslColor.Saturation, newColor.HslColor.Luminance);
                break;
            case eDrawingColorType.Preset:
                this.SetPresetColor(newColor.PresetColor.Color);
                break;
            case eDrawingColorType.System:
                this.SetSystemColor(newColor.SystemColor.Color);
                break;
            case eDrawingColorType.Scheme:
                this.SetSchemeColor(newColor.SchemeColor.Color);
                break;
        }
        //Variations should be added first, so temporary store the transforms and add the again
        List<IColorTransformItem>? trans = this.Transforms.Where(x=>((ISource)x)._fromStyleTemplate==false).ToList();
        this.Transforms.Clear();
        if (variation != null)
        {
            this.ApplyNewTransform(variation);
        }

        this.ApplyNewTransform(trans);
        this.ApplyNewTransform(newColor.Transforms, true);
    }

    private void ApplyNewTransform(IEnumerable<IColorTransformItem> transforms, bool isSourceStyleTemplate=false)
    {
        foreach (IColorTransformItem? t in transforms)
        {
            switch(t.Type)
            {
                case eColorTransformType.Alpha:
                    this.Transforms.AddAlpha(t.Value);
                    break;
                case eColorTransformType.AlphaMod:
                    this.Transforms.AddAlphaModulation(t.Value);
                    break;
                case eColorTransformType.AlphaOff:
                    this.Transforms.AddAlphaOffset(t.Value);
                    break;
                case eColorTransformType.Blue:
                    this.Transforms.AddBlue(t.Value);
                    break;
                case eColorTransformType.BlueMod:
                    this.Transforms.AddBlueModulation(t.Value);
                    break;
                case eColorTransformType.BlueOff:
                    this.Transforms.AddBlueOffset(t.Value);
                    break;
                case eColorTransformType.Comp:
                    this.Transforms.AddComplement();
                    break;
                case eColorTransformType.Gamma:
                    this.Transforms.AddGamma();
                    break;
                case eColorTransformType.Gray:
                    this.Transforms.AddGray();
                    break;
                case eColorTransformType.Green:
                    this.Transforms.AddGreen(t.Value);
                    break;
                case eColorTransformType.GreenMod:
                    this.Transforms.AddGreenModulation(t.Value);
                    break;
                case eColorTransformType.GreenOff:
                    this.Transforms.AddGreenOffset(t.Value);
                    break;
                case eColorTransformType.Hue:
                    this.Transforms.AddHue(t.Value);
                    break;
                case eColorTransformType.HueMod:
                    this.Transforms.AddHueModulation(t.Value);
                    break;
                case eColorTransformType.HueOff:
                    this.Transforms.AddHueOffset(t.Value);
                    break;
                case eColorTransformType.Inv:
                    this.Transforms.AddInverse();
                    break;
                case eColorTransformType.InvGamma:
                    this.Transforms.AddGamma();
                    break;
                case eColorTransformType.Lum:
                    this.Transforms.AddLuminance(t.Value);
                    break;
                case eColorTransformType.LumMod:
                    this.Transforms.AddLuminanceModulation(t.Value);
                    break;
                case eColorTransformType.LumOff:
                    this.Transforms.AddLuminanceOffset(t.Value);
                    break;
                case eColorTransformType.Red:
                    this.Transforms.AddRed(t.Value);
                    break;
                case eColorTransformType.RedMod:
                    this.Transforms.AddRedModulation(t.Value);
                    break;
                case eColorTransformType.RedOff:
                    this.Transforms.AddRedOffset(t.Value);
                    break;
                case eColorTransformType.Sat:
                    this.Transforms.AddSaturation(t.Value);
                    break;
                case eColorTransformType.SatMod:
                    this.Transforms.AddSaturationModulation(t.Value);
                    break;
                case eColorTransformType.SatOff:
                    this.Transforms.AddSaturationOffset(t.Value);
                    break;
                case eColorTransformType.Shade:
                    this.Transforms.AddShade(t.Value);
                    break;
                case eColorTransformType.Tint:
                    this.Transforms.AddTint(t.Value);
                    break;
            }
            if (isSourceStyleTemplate && this.Transforms.Count > 0)
            {
                ((ISource)this.Transforms.Last())._fromStyleTemplate = true;
            }
        }
    }
}