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

namespace OfficeOpenXml.Drawing.Theme;

/// <summary>
/// Defines the font scheme within the theme
/// </summary>
public class ExcelFontScheme : XmlHelper
{
    private ExcelPackage _pck;

    internal ExcelFontScheme(ExcelPackage pck, XmlNamespaceManager nameSpaceManager, XmlNode topNode)
        : base(nameSpaceManager, topNode) =>
        this._pck = pck;

    /// <summary>
    /// The name of the font scheme
    /// </summary>
    public string Name
    {
        get => this.GetXmlNodeString("@name");
        set => this.SetXmlNodeString("@name", value);
    }

    ExcelThemeFontCollection _majorFont;

    /// <summary>
    /// A collection of major fonts
    /// </summary>
    public ExcelThemeFontCollection MajorFont =>
        this._majorFont ??= new ExcelThemeFontCollection(this._pck,
                                                         this.NameSpaceManager,
                                                         this.TopNode.SelectSingleNode("a:majorFont", this.NameSpaceManager));

    ExcelThemeFontCollection _minorFont;

    /// <summary>
    /// A collection of minor fonts
    /// </summary>
    public ExcelThemeFontCollection MinorFont =>
        this._minorFont ??= new ExcelThemeFontCollection(this._pck,
                                                         this.NameSpaceManager,
                                                         this.TopNode.SelectSingleNode("a:minorFont", this.NameSpaceManager));
}