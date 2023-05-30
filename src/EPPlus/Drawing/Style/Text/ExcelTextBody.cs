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
using System.Xml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Properties for the textbody
/// </summary>
public class ExcelTextBody : XmlHelper
{
    private readonly string _path;

    internal ExcelTextBody(XmlNamespaceManager ns, XmlNode topNode, string path, string[] schemaNodeOrder = null)
        : base(ns, topNode)
    {
        this._path = path;

        this.AddSchemaNodeOrder(schemaNodeOrder,
                                new string[]
                                {
                                    "ln", "noFill", "solidFill", "gradFill", "pattFill", "blipFill", "latin", "ea", "cs", "sym", "hlinkClick",
                                    "hlinkMouseOver", "rtl", "extLst", "highlight", "kumimoji", "lang", "altLang", "sz", "b", "i", "u", "strike", "kern",
                                    "cap", "spc", "normalizeH", "baseline", "noProof", "dirty", "err", "smtClean", "smtId", "bmk"
                                });
    }

    /// <summary>
    /// The anchoring position within the shape
    /// </summary>
    public eTextAnchoringType Anchor
    {
        get => this.GetXmlNodeString($"{this._path}/@anchor").TranslateTextAchoring();
        set => this.SetXmlNodeString($"{this._path}/@anchor", value.TranslateTextAchoringText());
    }

    /// <summary>
    /// The centering of the text box.
    /// </summary>
    public bool AnchorCenter
    {
        get => this.GetXmlNodeBool($"{this._path}/@anchorCtr");
        set => this.SetXmlNodeBool($"{this._path}/@anchorCtr", value, false);
    }

    /// <summary>
    /// Underlined text
    /// </summary>
    public eUnderLineType UnderLine
    {
        get => this.GetXmlNodeString($"{this._path}/@u").TranslateUnderline();
        set
        {
            if (value == eUnderLineType.None)
            {
                this.DeleteNode($"{this._path}/@u");
            }
            else
            {
                this.SetXmlNodeString($"{this._path}/@u", value.TranslateUnderlineText());
            }
        }
    }

    /// <summary>
    /// The bottom inset of the bounding rectangle
    /// </summary>
    public double? BottomInsert
    {
        get => this.GetXmlNodeEmuToPtNull($"{this._path}/@bIns");
        set => this.SetXmlNodeEmuToPt($"{this._path}/@bIns", value);
    }

    /// <summary>
    /// The top inset of the bounding rectangle
    /// </summary>
    public double? TopInsert
    {
        get => this.GetXmlNodeEmuToPtNull($"{this._path}/@tIns");
        set => this.SetXmlNodeEmuToPt($"{this._path}/@tIns", value);
    }

    /// <summary>
    /// The right inset of the bounding rectangle
    /// </summary>
    public double? RightInsert
    {
        get => this.GetXmlNodeEmuToPtNull($"{this._path}/@rIns");
        set => this.SetXmlNodeEmuToPt($"{this._path}/@rIns", value);
    }

    /// <summary>
    /// The left inset of the bounding rectangle
    /// </summary>
    public double? LeftInsert
    {
        get => this.GetXmlNodeEmuToPtNull($"{this._path}/@lIns");
        set => this.SetXmlNodeEmuToPt($"{this._path}/@lIns", value);
    }

    /// <summary>
    /// The rotation that is being applied to the text within the bounding box
    /// </summary>
    public double? Rotation
    {
        get => this.GetXmlNodeAngel($"{this._path}/@rot");
        set => this.SetXmlNodeAngel($"{this._path}/@rot", value, "Rotation", -100000, 100000);
    }

    /// <summary>
    /// The space between text columns in the text area
    /// </summary>
    public double SpaceBetweenColumns
    {
        get => this.GetXmlNodeEmuToPt($"{this._path}/@spcCol");
        set
        {
            if (value < 0)
            {
                throw new ArgumentOutOfRangeException("SpaceBetweenColumns", "Can't be negative");
            }

            this.SetXmlNodeEmuToPt($"{this._path}/@spcCol", value);
        }
    }

    /// <summary>
    /// If the before and after paragraph spacing defined by the user is to be respected
    /// </summary>
    public bool ParagraphSpacing
    {
        get => this.GetXmlNodeBool($"{this._path}/@spcFirstLastPara");
        set => this.SetXmlNodeBool($"{this._path}/@spcFirstLastPara", value);
    }

    /// <summary>
    /// 
    /// </summary>
    public bool TextUpright
    {
        get => this.GetXmlNodeBool($"{this._path}/@upright");
        set => this.SetXmlNodeBool($"{this._path}/@upright", value);
    }

    /// <summary>
    /// If the line spacing is decided in a simplistic manner using the font scene
    /// </summary>
    public bool CompatibleLineSpacing
    {
        get => this.GetXmlNodeBool($"{this._path}/@compatLnSpc");
        set => this.SetXmlNodeBool($"{this._path}/@compatLnSpc", value);
    }

    /// <summary>
    /// Forces the text to be rendered anti-aliased
    /// </summary>
    public bool ForceAntiAlias
    {
        get => this.GetXmlNodeBool($"{this._path}/@forceAA");
        set => this.SetXmlNodeBool($"{this._path}/@forceAA", value);
    }

    /// <summary>
    /// If the text within this textbox is converted from a WordArt object.
    /// </summary>
    public bool FromWordArt
    {
        get => this.GetXmlNodeBool($"{this._path}/@fromWordArt");
        set => this.SetXmlNodeBool($"{this._path}/@fromWordArt", value);
    }

    /// <summary>
    /// If the text should be displayed vertically
    /// </summary>
    public eTextVerticalType VerticalText
    {
        get => this.GetXmlNodeString($"{this._path}/@vert").TranslateTextVertical();
        set => this.SetXmlNodeString($"{this._path}/@vert", value.TranslateTextVerticalText());
    }

    /// <summary>
    /// If the text can flow out horizontaly
    /// </summary>
    public eTextHorizontalOverflow HorizontalTextOverflow
    {
        get => this.GetXmlNodeString($"{this._path}/@horzOverflow").ToEnum(eTextHorizontalOverflow.Overflow);
        set => this.SetXmlNodeString($"{this._path}/@horzOverflow", value.ToEnumString());
    }

    /// <summary>
    /// If the text can flow out of the bounding box vertically
    /// </summary>
    public eTextVerticalOverflow VerticalTextOverflow
    {
        get => this.GetXmlNodeString($"{this._path}/@vertOverflow").ToEnum(eTextVerticalOverflow.Overflow);
        set => this.SetXmlNodeString($"{this._path}/@vertOverflow", value.ToEnumString());
    }

    /// <summary>
    /// How text is wrapped
    /// </summary>
    public eTextWrappingType WrapText
    {
        get => this.GetXmlNodeString($"{this._path}/@wrap").ToEnum(eTextWrappingType.Square);
        set => this.SetXmlNodeString($"{this._path}/@wrap", value.ToEnumString());
    }

    /// <summary>
    /// The text within the text body should be normally auto-fited
    /// </summary>
    public eTextAutofit TextAutofit
    {
        get
        {
            if (this.ExistsNode($"{this._path}/a:normAutofit"))
            {
                return eTextAutofit.NormalAutofit;
            }
            else if (this.ExistsNode($"{this._path}/a:spAutoFit"))
            {
                return eTextAutofit.ShapeAutofit;
            }
            else
            {
                return eTextAutofit.NoAutofit;
            }
        }
        set
        {
            switch (value)
            {
                case eTextAutofit.NormalAutofit:
                    if (value == this.TextAutofit)
                    {
                        return;
                    }

                    this.DeleteNode($"{this._path}/a:spAutoFit");
                    this.DeleteNode($"{this._path}/a:noAutofit");
                    _ = this.CreateNode($"{this._path}/a:normAutofit");

                    break;

                case eTextAutofit.ShapeAutofit:
                    this.DeleteNode($"{this._path}/a:noAutofit");
                    this.DeleteNode($"{this._path}/a:normAutofit");
                    _ = this.CreateNode($"{this._path}/a:spAutofit");

                    break;

                case eTextAutofit.NoAutofit:
                    this.DeleteNode($"{this._path}/a:spAutoFit");
                    this.DeleteNode($"{this._path}/a:normAutofit");
                    _ = this.CreateNode($"{this._path}/a:noAutofit");

                    break;
            }
        }
    }

    /// <summary>
    /// The percentage of the original font size to which each run in the text body is scaled.
    /// This propery only applies when the TextAutofit property is set to NormalAutofit
    /// </summary>
    public double? AutofitNormalFontScale
    {
        get => this.GetXmlNodePercentage($"{this._path}/a:normAutofit/@fontScale");
        set
        {
            if (this.TextAutofit != eTextAutofit.NormalAutofit)
            {
                throw new ArgumentException("AutofitNormalFontScale", "TextAutofit must be set to NormalAutofit to use set this property");
            }

            this.SetXmlNodePercentage($"{this._path}/a:normAutofit/@fontScale", value, false);
        }
    }

    /// <summary>
    /// The percentage by which the line spacing of each paragraph is reduced.
    /// This propery only applies when the TextAutofit property is set to NormalAutofit
    /// </summary>
    public double? LineSpaceReduction
    {
        get => this.GetXmlNodePercentage($"{this._path}/a:normAutofit/@lnSpcReduction");
        set
        {
            if (this.TextAutofit != eTextAutofit.NormalAutofit)
            {
                throw new ArgumentException("LineSpaceReduction", "TextAutofit must be set to NormalAutofit to use set this property");
            }

            this.SetXmlNodePercentage($"{this._path}/a:normAutofit/@lnSpcReduction", value, false);
        }
    }

    internal XmlElement PathElement
    {
        get
        {
            XmlElement? node = (XmlElement)this.GetNode(this._path);

            if (node == null)
            {
                return (XmlElement)this.CreateNode(this._path);
            }
            else
            {
                return node;
            }
        }
    }

    internal void SetFromXml(XmlElement copyFromElement)
    {
        XmlElement? element = this.PathElement;

        foreach (XmlAttribute a in copyFromElement.Attributes)
        {
            _ = element.SetAttribute(a.Name, a.NamespaceURI, a.Value);
        }
    }
}