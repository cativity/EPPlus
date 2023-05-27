/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/21/2020         EPPlus Software AB           EPPlus 5.5 
 *************************************************************************************************/

using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls;

/// <summary>
/// Represents a button form control
/// </summary>
public class ExcelControlButton : ExcelControlWithText
{
    internal ExcelControlButton(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent = null)
        : base(drawings, drawNode, name, parent)
    {
        this.SetSize(90, 30); //Default size
    }

    internal ExcelControlButton(ExcelDrawings drawings,
                                XmlNode drawNode,
                                ControlInternal control,
                                ZipPackagePart part,
                                XmlDocument controlPropertiesXml,
                                ExcelGroupShape parent = null)
        : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
    {
    }

    /// <summary>
    /// The type of form control
    /// </summary>
    public override eControlType ControlType => eControlType.Button;

    private ExcelControlMargin _margin;

    /// <summary>
    /// The buttons margin settings
    /// </summary>
    public ExcelControlMargin Margin
    {
        get { return this._margin ??= new ExcelControlMargin(this); }
    }

    /// <summary>
    /// The buttons text layout flow
    /// </summary>
    public eLayoutFlow LayoutFlow { get; set; }

    /// <summary>
    /// Text orientation
    /// </summary>
    public eShapeOrientation Orientation { get; set; }

    /// <summary>
    /// The reading order for the text
    /// </summary>
    public eReadingOrder ReadingOrder { get; set; }

    /// <summary>
    /// If size is automatic
    /// </summary>
    public bool AutomaticSize { get; set; }

    /// <summary>
    /// Text Anchoring for the text body
    /// </summary>
    internal eTextAnchoringType TextAnchor
    {
        get { return this.TextBody.Anchor; }
        set { this.TextBody.Anchor = value; }
    }

    private string _textAlignPath = "xdr:sp/xdr:txBody/a:p/a:pPr/@algn";

    /// <summary>
    /// How the text is aligned
    /// </summary>
    public eTextAlignment TextAlignment
    {
        get
        {
            switch (this.GetXmlNodeString(this._textAlignPath))
            {
                case "ctr":
                    return eTextAlignment.Center;

                case "r":
                    return eTextAlignment.Right;

                case "dist":
                    return eTextAlignment.Distributed;

                case "just":
                    return eTextAlignment.Justified;

                case "justLow":
                    return eTextAlignment.JustifiedLow;

                case "thaiDist":
                    return eTextAlignment.ThaiDistributed;

                default:
                    return eTextAlignment.Left;
            }
        }
        set
        {
            switch (value)
            {
                case eTextAlignment.Right:
                    this.SetXmlNodeString(this._textAlignPath, "r");

                    break;

                case eTextAlignment.Center:
                    this.SetXmlNodeString(this._textAlignPath, "ctr");

                    break;

                case eTextAlignment.Distributed:
                    this.SetXmlNodeString(this._textAlignPath, "dist");

                    break;

                case eTextAlignment.Justified:
                    this.SetXmlNodeString(this._textAlignPath, "just");

                    break;

                case eTextAlignment.JustifiedLow:
                    this.SetXmlNodeString(this._textAlignPath, "justLow");

                    break;

                case eTextAlignment.ThaiDistributed:
                    this.SetXmlNodeString(this._textAlignPath, "thaiDist");

                    break;

                default:
                    this.DeleteNode(this._textAlignPath);

                    break;
            }
        }
    }

    internal override void UpdateXml()
    {
        base.UpdateXml();
        this.Margin.UpdateXml();
        XmlHelper? vmlHelper = XmlHelperFactory.Create(this._vmlProp.NameSpaceManager, this._vmlProp.TopNode.ParentNode);
        string? style = "layout-flow:" + this.LayoutFlow.TranslateString() + ";mso-layout-flow-alt:" + this.Orientation.TranslateString();

        if (this.ReadingOrder == eReadingOrder.RightToLeft)
        {
            style += ";direction:RTL";
        }
        else if (this.ReadingOrder == eReadingOrder.ContextDependent)
        {
            style += ";mso-direction-alt:auto";
        }

        if (this.AutomaticSize)
        {
            style += ";mso-fit-shape-to-text:t";
        }

        vmlHelper.SetXmlNodeString("v:textbox/@style", style);
    }
}