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

using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect;

/// <summary>
/// The shadow effect applied to a drawing
/// </summary>
public abstract class ExcelDrawingShadowEffect : ExcelDrawingShadowEffectBase
{
    private readonly string _directionPath = "{0}/@dir";

    internal ExcelDrawingShadowEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path)
        : base(nameSpaceManager, topNode, schemaNodeOrder, path) =>
        this._directionPath = string.Format(this._directionPath, path);

    ExcelDrawingColorManager _color;

    /// <summary>
    /// The color of the shadow effect
    /// </summary>
    public ExcelDrawingColorManager Color => this._color ??= new ExcelDrawingColorManager(this.NameSpaceManager, this.TopNode, this._path, this.SchemaNodeOrder);

    /// <summary>
    /// The direction angle to offset the shadow.
    /// Ranges from 0 to 360
    /// </summary>
    public double? Direction
    {
        get => this.GetXmlNodeAngel(this._directionPath);
        set
        {
            this.InitXml();
            this.SetXmlNodeAngel(this._directionPath, value, "Direction");
        }
    }

    /// <summary>
    /// Inizialize the xml
    /// </summary>
    protected internal void InitXml()
    {
        if (this._color == null)
        {
            this.Color.SetPresetColor(ePresetColor.Black);
            this.Color.Transforms.AddAlpha(50);
        }
    }
}