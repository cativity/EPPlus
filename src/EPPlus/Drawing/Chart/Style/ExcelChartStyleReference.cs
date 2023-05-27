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

namespace OfficeOpenXml.Drawing.Chart.Style;

/// <summary>
/// A reference from a chart style to the theme collection
/// </summary>
public class ExcelChartStyleReference : XmlHelper
{
    string _path;

    internal ExcelChartStyleReference(XmlNamespaceManager nsm, XmlNode topNode, string path)
        : base(nsm, topNode)
    {
        this._path = path;
    }

    /// <summary>
    /// The index to the theme style matrix.
    /// <seealso cref="ExcelWorkbook.ThemeManager"/>
    /// </summary>
    public int Index
    {
        get { return this.GetXmlNodeInt($"{this._path}/@idx"); }
        set
        {
            if (value < 0)
            {
                throw new ArgumentOutOfRangeException("Index", "Can't be negative");
            }

            this.SetXmlNodeString($"{this._path}/@idx", value.ToString(CultureInfo.InvariantCulture));
        }
    }

    ExcelChartStyleColorManager _color = null;

    /// <summary>
    /// The color to be used for the reference. 
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