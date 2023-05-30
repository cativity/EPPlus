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
using OfficeOpenXml.Drawing.Style.Fill;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme;

/// <summary>
/// Linestyle for a theme
/// </summary>
public class ExcelThemeLine : XmlHelper
{
    internal ExcelThemeLine(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
        : base(nameSpaceManager, topNode)
    {
        this.SchemaNodeOrder = new string[] { "noFill", "solidFill", "gradientFill", "pattFill", "prstDash", "round", "bevel", "miter", "headEnd", " tailEnd" };
    }

    const string widthPath = "@w";

    /// <summary>
    /// Line width, in EMU's
    /// 
    /// 1 Pixel      =   9525
    /// 1 Pt         =   12700
    /// 1 cm         =   360000 
    /// 1 US inch    =   914400
    /// </summary>
    public int Width
    {
        get { return this.GetXmlNodeInt(widthPath); }
        set { this.SetXmlNodeString(widthPath, value.ToString(CultureInfo.InvariantCulture)); }
    }

    const string CapPath = "@cap";

    /// <summary>
    /// The ending caps for the line
    /// </summary>
    public eLineCap Cap
    {
        get { return EnumTransl.ToLineCap(this.GetXmlNodeString(CapPath)); }
        set { this.SetXmlNodeString(CapPath, EnumTransl.FromLineCap(value)); }
    }

    const string CompoundPath = "@cmpd";

    /// <summary>
    /// The compound line type to be used for the underline stroke
    /// </summary>
    public eCompundLineStyle CompoundLineStyle
    {
        get { return EnumTransl.ToLineCompound(this.GetXmlNodeString(CompoundPath)); }
        set { this.SetXmlNodeString(CompoundPath, EnumTransl.FromLineCompound(value)); }
    }

    const string PenAlignmentPath = "@algn";

    /// <summary>
    /// Specifies the pen alignment type for use within a text body
    /// </summary>
    public ePenAlignment Alignment
    {
        get { return EnumTransl.ToPenAlignment(this.GetXmlNodeString(PenAlignmentPath)); }
        set { this.SetXmlNodeString(PenAlignmentPath, EnumTransl.FromPenAlignment(value)); }
    }

    ExcelDrawingFill _fill;

    /// <summary>
    /// Access to fill properties
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            if (this._fill == null)
            {
                if (!(this.TopNode.HasChildNodes && this.TopNode.ChildNodes[0].LocalName.EndsWith("Fill")))
                {
                    this._fill = new ExcelDrawingFill(null, this.NameSpaceManager, this.TopNode.ChildNodes[0], "", this.SchemaNodeOrder);
                }
                else
                {
                    _ = this.CreateNode("a:solidFill");
                    this._fill = new ExcelDrawingFill(null, this.NameSpaceManager, this.TopNode.ChildNodes[0], "", this.SchemaNodeOrder);
                    this.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Style);
                }
            }

            return this._fill;
        }
    }

    const string StylePath = "a:prstDash/@val";

    /// <summary>
    /// Preset line dash
    /// </summary>
    public eLineStyle Style
    {
        get { return EnumTransl.ToLineStyle(this.GetXmlNodeString(StylePath)); }
        set { this.SetXmlNodeString(StylePath, EnumTransl.FromLineStyle(value)); }
    }

    const string BevelPath = "a:bevel";
    const string RoundPath = "a:round";
    const string MiterPath = "a:miter";

    /// <summary>
    /// The shape that lines joined together have
    /// </summary>
    public eLineJoin? Join
    {
        get
        {
            if (this.ExistsNode(BevelPath))
            {
                return eLineJoin.Bevel;
            }
            else if (this.ExistsNode(RoundPath))
            {
                return eLineJoin.Round;
            }
            else if (this.ExistsNode(MiterPath))
            {
                return eLineJoin.Miter;
            }
            else
            {
                return null;
            }
        }
        set
        {
            if (value == eLineJoin.Bevel)
            {
                _ = this.CreateNode(BevelPath);
                this.DeleteNode(RoundPath);
                this.DeleteNode(MiterPath);
            }
            else if (value == eLineJoin.Round)
            {
                _ = this.CreateNode(RoundPath);
                this.DeleteNode(BevelPath);
                this.DeleteNode(MiterPath);
            }
            else
            {
                _ = this.CreateNode(MiterPath);
                this.DeleteNode(RoundPath);
                this.DeleteNode(BevelPath);
            }
        }
    }

    const string MiterJoinLimitPath = "a:miter/@lim";

    /// <summary>
    /// How much lines are extended to form a miter join
    /// </summary>
    public double? MiterJoinLimit
    {
        get { return this.GetXmlNodePercentage(MiterJoinLimitPath); }
        set
        {
            this.Join = eLineJoin.Miter;
            this.SetXmlNodePercentage(MiterJoinLimitPath, value);
        }
    }

    ExcelDrawingLineEnd _headEnd = null;

    /// <summary>
    /// Properties for drawing line head ends
    /// </summary>
    public ExcelDrawingLineEnd HeadEnd
    {
        get
        {
            if (this._headEnd == null)
            {
                return new ExcelDrawingLineEnd(this.NameSpaceManager, this.TopNode, "a:headEnd", Init);
            }

            return this._headEnd;
        }
    }

    ExcelDrawingLineEnd _tailEnd = null;

    /// <summary>
    /// Properties for drawing line tail ends
    /// </summary>
    public ExcelDrawingLineEnd TailEnd
    {
        get
        {
            if (this._tailEnd == null)
            {
                return new ExcelDrawingLineEnd(this.NameSpaceManager, this.TopNode, "a:tailEnd", Init);
            }

            return this._tailEnd;
        }
    }

    internal XmlElement LineElement
    {
        get { return this.TopNode as XmlElement; }
    }

    private static void Init()
    {
    }
}