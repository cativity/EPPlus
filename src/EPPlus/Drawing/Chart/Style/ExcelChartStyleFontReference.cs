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

using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.Extensions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style;

/// <summary>
/// A reference to a theme font collection from the chart style manager
/// </summary>
public class ExcelChartStyleFontReference : XmlHelper
{
    string _path;

    internal ExcelChartStyleFontReference(XmlNamespaceManager nsm, XmlNode topNode, string path)
        : base(nsm, topNode)
    {
        this._path = path;
    }

    /// <summary>
    /// The index to the style matrix.
    /// This property referes to the theme
    /// </summary>
    public eThemeFontCollectionType Index
    {
        get { return this.GetXmlNodeString($"{this._path}/@idx").ToEnum(eThemeFontCollectionType.None); }
        set { this.SetXmlNodeString($"{this._path}/@idx", value.ToEnumString()); }
    }

    ExcelChartStyleColorManager _color = null;

    /// <summary>
    /// The color of the font
    /// This will replace any the StyleClr node in the chart style xml.
    /// </summary>
    public ExcelChartStyleColorManager Color
    {
        get { return this._color ??= new ExcelChartStyleColorManager(this.NameSpaceManager, this.TopNode, this._path, this.SchemaNodeOrder); }
    }

    /// <summary>
    /// If the reference has a color
    /// </summary>
    public bool HasColor
    {
        get
        {
            XmlNode? node = this.GetNode(this._path);

            return node != null && node.HasChildNodes;
        }
    }
}