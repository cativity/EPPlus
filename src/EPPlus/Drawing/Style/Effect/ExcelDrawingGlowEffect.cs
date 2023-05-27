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

using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect;

/// <summary>
/// The glow effect, in which a color blurred outline is added outside the edges of the drawing
/// </summary>
public class ExcelDrawingGlowEffect : ExcelDrawingEffectBase
{
    private readonly string _radiusPath = "{0}/@rad";

    internal ExcelDrawingGlowEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path)
        : base(nameSpaceManager, topNode, schemaNodeOrder, path)
    {
        this._radiusPath = string.Format(this._radiusPath, path);
    }

    ExcelDrawingColorManager _color = null;

    /// <summary>
    /// The color of the glow
    /// </summary>
    public ExcelDrawingColorManager Color
    {
        get { return this._color ??= new ExcelDrawingColorManager(this.NameSpaceManager, this.TopNode, this._path, this.SchemaNodeOrder); }
    }

    /// <summary>
    /// The radius of the glow in pixels
    /// </summary>
    public double? Radius
    {
        get { return this.GetXmlNodeEmuToPtNull(this._radiusPath) ?? 0; }
        set
        {
            this.SetXmlNodeEmuToPt(this._radiusPath, value);
            this.InitXml();
        }
    }

    private void InitXml()
    {
        if (this._color == null)
        {
            this.Color.SetPresetColor(ePresetColor.Black);
            this.Color.Transforms.AddAlpha(50);
        }
    }
}