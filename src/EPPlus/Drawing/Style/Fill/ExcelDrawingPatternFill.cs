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
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill;

/// <summary>
/// A pattern fill.
/// </summary>
public class ExcelDrawingPatternFill : ExcelDrawingFillBase
{
    string[] _schemaNodeOrder;
    internal ExcelDrawingPatternFill(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string fillPath, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode, fillPath, initXml)
    {
        this._schemaNodeOrder = XmlHelper.CopyToSchemaNodeOrder(schemaNodeOrder, new string[] { "fgClr", "bgClr" });
        this.GetXml();
    }
    /// <summary>
    /// The fillstyle, always PatternFill
    /// </summary>
    public override eFillStyle Style
    {
        get
        {
            return eFillStyle.PatternFill;
        }
    }
    private eFillPatternStyle _pattern;
    /// <summary>
    /// The preset pattern to use
    /// </summary>
    public eFillPatternStyle PatternType
    {
        get
        {
            return this._pattern;
        }
        set
        {
            this._pattern = value;
        }
    }
    ExcelDrawingColorManager _fgColor = null;
    /// <summary>
    /// Foreground color
    /// </summary>
    public ExcelDrawingColorManager ForegroundColor
    {
        get
        {
            return this._fgColor ??= new ExcelDrawingColorManager(this._nsm, this._topNode, "a:fgClr", this._schemaNodeOrder, this._initXml);
        }
    }
    ExcelDrawingColorManager _bgColor = null;
    /// <summary>
    /// Background color
    /// </summary>
    public ExcelDrawingColorManager BackgroundColor
    {
        get
        {
            return this._bgColor ??= new ExcelDrawingColorManager(this._nsm, this._topNode, "a:bgClr", this._schemaNodeOrder, this._initXml);
        }
    }


    internal override string NodeName
    {
        get
        {
            return "a:patternFill";
        }
    }

    internal override void SetXml(XmlNamespaceManager nsm, XmlNode node)
    {
        this._initXml?.Invoke();
        if (this._xml == null)
        {
            if(string.IsNullOrEmpty(this._fillPath))
            {
                this.InitXml(nsm, node,"");
            }
            else
            {
                this.CreateXmlHelper();
            }
        }

        this._xml.SetXmlNodeString("@prst", this.PatternType.ToEnumString());
        XmlNode? fgNode= this._xml.CreateNode("a:fgClr");
        ExcelDrawingThemeColorManager.SetXml(nsm, fgNode);

        XmlNode? bgNode = this._xml.CreateNode("a:bgClr");
        ExcelDrawingThemeColorManager.SetXml(nsm, bgNode);
    }
    internal override void GetXml()
    {
        this.PatternType = this._xml.GetXmlNodeString("@prst").ToEnum(eFillPatternStyle.Pct5);
    }

    internal override void UpdateXml()
    {
        if (this._xml == null)
        {
            this.CreateXmlHelper();
        }

        this.SetXml(this._nsm, this._xml.TopNode);
    }
}