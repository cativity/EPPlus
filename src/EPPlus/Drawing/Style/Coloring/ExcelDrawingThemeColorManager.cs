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

using System.Drawing;
using System.Xml;
using System;

namespace OfficeOpenXml.Drawing.Style.Coloring;

/// <summary>
/// Manages colors in a theme 
/// </summary>
public class ExcelDrawingThemeColorManager
{
    /// <summary>
    /// Namespace manager
    /// </summary>
    internal protected XmlNamespaceManager _nameSpaceManager;

    /// <summary>
    /// The top node
    /// </summary>
    internal protected XmlNode _topNode;

    /// <summary>
    /// The node of the supplied path
    /// </summary>
    internal protected XmlNode _pathNode;

    /// <summary>
    /// The node of the color object
    /// </summary>
    internal protected XmlNode _colorNode;

    /// <summary>
    /// Init method
    /// </summary>
    internal protected Action _initMethod;

    /// <summary>
    /// The x-path
    /// </summary>
    internal protected string _path;

    /// <summary>
    /// Order of the elements according to the xml schema
    /// </summary>
    internal protected string[] _schemaNodeOrder;

    internal ExcelDrawingThemeColorManager(XmlNamespaceManager nameSpaceManager,
                                           XmlNode topNode,
                                           string path,
                                           string[] schemaNodeOrder,
                                           Action initMethod = null)
    {
        this._nameSpaceManager = nameSpaceManager;
        this._topNode = topNode;
        this._path = path;
        this._initMethod = initMethod;
        this._pathNode = this.GetPathNode();
        this._schemaNodeOrder = schemaNodeOrder;

        if (this._pathNode == null)
        {
            return;
        }

        if (IsTopNodeColorNode(this._topNode))
        {
            this._colorNode = this._pathNode;
        }
        else
        {
            this._colorNode = this._pathNode.FirstChild;
        }

        if (this._colorNode == null)
        {
            return;
        }

        switch (this._colorNode.LocalName)
        {
            case "sysClr":
                this.ColorType = eDrawingColorType.System;
                this.SystemColor = new ExcelDrawingSystemColor(this._nameSpaceManager, this._pathNode.FirstChild);

                break;

            case "scrgbClr":
                this.ColorType = eDrawingColorType.RgbPercentage;
                this.RgbPercentageColor = new ExcelDrawingRgbPercentageColor(this._nameSpaceManager, this._pathNode.FirstChild);

                break;

            case "hslClr":
                this.ColorType = eDrawingColorType.Hsl;
                this.HslColor = new ExcelDrawingHslColor(this._nameSpaceManager, this._pathNode.FirstChild);

                break;

            case "prstClr":
                this.ColorType = eDrawingColorType.Preset;
                this.PresetColor = new ExcelDrawingPresetColor(this._nameSpaceManager, this._pathNode.FirstChild);

                break;

            case "srgbClr":
                this.ColorType = eDrawingColorType.Rgb;
                this.RgbColor = new ExcelDrawingRgbColor(this._nameSpaceManager, this._pathNode.FirstChild);

                break;

            default:
                this.ColorType = eDrawingColorType.None;

                break;
        }
    }

    private static bool IsTopNodeColorNode(XmlNode topNode)
    {
        return topNode.LocalName.EndsWith("Clr");
    }

    /// <summary>
    /// The type of color.
    /// Each type has it's own property and set-method.       
    /// <see cref="SetRgbColor(Color, bool)"/>
    /// <see cref="SetRgbPercentageColor(double, double, double)"/>
    /// <see cref="SetHslColor(double, double, double)" />
    /// <see cref="SetPresetColor(Color)"/>
    /// <see cref="SetPresetColor(ePresetColor)"/>
    /// <see cref="SetSystemColor(eSystemColor)"/>
    /// <see cref="ExcelDrawingColorManager.SetSchemeColor(eSchemeColor)"/>
    /// </summary>
    public eDrawingColorType ColorType { get; internal protected set; } = eDrawingColorType.None;

    internal static void SetXml(XmlNamespaceManager nameSpaceManager, XmlNode node)
    {
    }

    ExcelColorTransformCollection _transforms;

    /// <summary>
    /// Color transformations
    /// </summary>
    public ExcelColorTransformCollection Transforms
    {
        get
        {
            if (this.ColorType == eDrawingColorType.None)
            {
                return null;
            }

            return this._transforms ??= new ExcelColorTransformCollection(this._nameSpaceManager, this._colorNode);
        }
    }

    /// <summary>
    /// A rgb color.
    /// This property has a value when Type is set to Rgb
    /// </summary>
    public ExcelDrawingRgbColor RgbColor { get; private set; }

    /// <summary>
    /// A rgb precentage color.
    /// This property has a value when Type is set to RgbPercentage
    /// </summary>
    public ExcelDrawingRgbPercentageColor RgbPercentageColor { get; private set; }

    /// <summary>
    /// A hsl color.
    /// This property has a value when Type is set to Hsl
    /// </summary>
    public ExcelDrawingHslColor HslColor { get; private set; }

    /// <summary>
    /// A preset color.
    /// This property has a value when Type is set to Preset
    /// </summary>
    public ExcelDrawingPresetColor PresetColor { get; private set; }

    /// <summary>
    /// A system color.
    /// This property has a value when Type is set to System
    /// </summary>
    public ExcelDrawingSystemColor SystemColor { get; private set; }

    /// <summary>
    /// Sets a rgb color.
    /// </summary>
    /// <param name="color">The color</param>
    /// <param name="setAlpha">Apply the alpha part of the Color to the <see cref="Transforms"/> collection</param>
    public void SetRgbColor(Color color, bool setAlpha = false)
    {
        this.ColorType = eDrawingColorType.Rgb;
        this.ResetColors(ExcelDrawingRgbColor.NodeName);

        if (setAlpha && color.A != 0xFF)
        {
            this.Transforms.RemoveOfType(eColorTransformType.Alpha);
            this.Transforms.AddAlpha(((double)color.A + 1) / 256 * 100);
        }

        this.RgbColor = new ExcelDrawingRgbColor(this._nameSpaceManager, this._colorNode) { Color = color };
    }

    /// <summary>
    /// Sets a rgb precentage color
    /// </summary>
    /// <param name="redPercentage">Red percentage</param>
    /// <param name="greenPercentage">Green percentage</param>
    /// <param name="bluePercentage">Bluepercentage</param>
    public void SetRgbPercentageColor(double redPercentage, double greenPercentage, double bluePercentage)
    {
        this.ColorType = eDrawingColorType.RgbPercentage;
        this.ResetColors(ExcelDrawingRgbPercentageColor.NodeName);

        this.RgbPercentageColor = new ExcelDrawingRgbPercentageColor(this._nameSpaceManager, this._colorNode)
        {
            RedPercentage = redPercentage, GreenPercentage = greenPercentage, BluePercentage = bluePercentage
        };
    }

    /// <summary>
    /// Sets a hsl color
    /// </summary>
    /// <param name="hue">The hue angle. From 0-360</param>
    /// <param name="saturation">The saturation percentage. From 0-100</param>
    /// <param name="luminance">The luminance percentage. From 0-100</param>
    public void SetHslColor(double hue, double saturation, double luminance)
    {
        this.ColorType = eDrawingColorType.Hsl;
        this.ResetColors(ExcelDrawingHslColor.NodeName);
        this.HslColor = new ExcelDrawingHslColor(this._nameSpaceManager, this._colorNode) { Hue = hue, Saturation = saturation, Luminance = luminance };
    }

    /// <summary>
    /// Sets a preset color.
    /// Must be a named color. Can't be color.Empty.
    /// </summary>
    /// <param name="color">Color</param>
    public void SetPresetColor(Color color)
    {
        this.ColorType = eDrawingColorType.Preset;
        this.ResetColors(ExcelDrawingPresetColor.NodeName);
        this.PresetColor = new ExcelDrawingPresetColor(this._nameSpaceManager, this._colorNode) { Color = ExcelDrawingPresetColor.GetPresetColor(color) };
    }

    /// <summary>
    /// Sets a preset color.
    /// </summary>
    /// <param name="presetColor">The color</param>
    public void SetPresetColor(ePresetColor presetColor)
    {
        this.ColorType = eDrawingColorType.Preset;
        this.ResetColors(ExcelDrawingPresetColor.NodeName);
        this.PresetColor = new ExcelDrawingPresetColor(this._nameSpaceManager, this._colorNode) { Color = presetColor };
    }

    /// <summary>
    /// Sets a system color
    /// </summary>
    /// <param name="systemColor">The colors</param>
    public void SetSystemColor(eSystemColor systemColor)
    {
        this.ColorType = eDrawingColorType.System;
        this.ResetColors(ExcelDrawingSystemColor.NodeName);
        this.SystemColor = new ExcelDrawingSystemColor(this._nameSpaceManager, this._colorNode) { Color = systemColor };
    }

    /// <summary>
    /// Reset the color objects
    /// </summary>
    /// <param name="newNodeName">The new color node name</param>
    internal protected virtual void ResetColors(string newNodeName)
    {
        if (this._colorNode == null)
        {
            XmlHelper? xml = XmlHelperFactory.Create(this._nameSpaceManager, this._topNode);
            xml.SchemaNodeOrder = this._schemaNodeOrder;
            string? colorPath = string.IsNullOrEmpty(this._path) ? newNodeName : this._path + "/" + newNodeName;
            this._colorNode = xml.CreateNode(colorPath);
            this._initMethod?.Invoke();
        }

        if (this._colorNode.Name == newNodeName)
        {
            return;
        }
        else
        {
            this._transforms = null;
            this.ChangeType(newNodeName);
        }

        this.RgbColor = null;
        this.RgbPercentageColor = null;
        this.HslColor = null;
        this.PresetColor = null;
        this.SystemColor = null;
    }

    private void ChangeType(string type)
    {
        if (this._topNode == this._colorNode)
        {
            XmlHelper? xh = XmlHelperFactory.Create(this._nameSpaceManager, this._topNode);
            _ = xh.ReplaceElement(this._colorNode, type);
        }
        else
        {
            XmlNode? p = this._colorNode.ParentNode;
            p.InnerXml = $"<{type} />";
            this._colorNode = p.FirstChild;
        }
    }

    private XmlNode GetPathNode()
    {
        if (this._pathNode == null)
        {
            if (string.IsNullOrEmpty(this._path))
            {
                this._pathNode = this._topNode;
            }
            else
            {
                this._pathNode = this._topNode.SelectSingleNode(this._path, this._nameSpaceManager);
            }
        }

        return this._pathNode;
    }
}