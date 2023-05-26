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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill;

/// <summary>
/// A solid fill.
/// </summary>
public class ExcelDrawingSolidFill : ExcelDrawingFillBase
{
    string[] _schemaNodeOrder;
    internal ExcelDrawingSolidFill(XmlNamespaceManager nsm, XmlNode topNode, string fillPath, string[]  schemaNodeOrder, Action initXml) : base(nsm, topNode, fillPath, initXml)
    {
        this._schemaNodeOrder = schemaNodeOrder;
        this.GetXml();
    }
    /// <summary>
    /// The fill style
    /// </summary>
    public override eFillStyle Style
    {
        get
        {
            return eFillStyle.SolidFill;
        }
    }
    ExcelDrawingColorManager _color = null;

    /// <summary>
    /// The color of the fill
    /// </summary>
    public ExcelDrawingColorManager Color
    {
        get
        {
            return this._color ??= new ExcelDrawingColorManager(this._nsm, this._topNode, this._fillPath, this._schemaNodeOrder, this._initXml);
        }
    }

    internal override string NodeName
    {
        get
        {
            return "a:solidFill";
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

        this.CheckTypeChange(this.NodeName);
        if(this._color==null)
        {
            this.Color.SetPresetColor(ePresetColor.Black);
        }
        ExcelDrawingThemeColorManager.SetXml(nsm, node);
    }
    internal override void GetXml()
    {
            
    }
    internal override void UpdateXml()
    {
    }
}