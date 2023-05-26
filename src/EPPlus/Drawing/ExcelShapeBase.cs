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
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Base class for drawing-shape objects
/// </summary>
public class ExcelShapeBase : ExcelDrawing
{
    internal string _shapeStylePath = "{0}xdr:spPr/a:prstGeom/@prst";
    private string _fillPath = "{0}xdr:spPr";
    private string _borderPath = "{0}xdr:spPr/a:ln";
    private string _effectPath = "{0}xdr:spPr/a:effectLst";
    private string _headEndPath = "{0}xdr:spPr/a:ln/a:headEnd";
    private string _tailEndPath = "{0}xdr:spPr/a:ln/a:tailEnd";
    private string _textPath = "{0}xdr:txBody/a:p/a:r/a:t";
    private string _lockTextPath = "{0}@fLocksText";
    private string _textAnchoringPath = "{0}xdr:txBody/a:bodyPr/@anchor";
    private string _textAnchoringCtlPath = "{0}xdr:txBody/a:bodyPr/@anchorCtr";
    private string _paragraphPath = "{0}xdr:txBody/a:p";
    private string _textAlignPath = "{0}xdr:txBody/a:p/a:pPr/@algn";
    private string _indentAlignPath = "{0}xdr:txBody/a:p/a:pPr/@lvl";
    private string _textVerticalPath = "{0}xdr:txBody/a:bodyPr/@vert";
    private string _fontPath = "{0}xdr:txBody/a:p/a:pPr/a:defRPr";
    private string _textBodyPath = "{0}xdr:txBody/a:bodyPr";
    internal ExcelShapeBase(ExcelDrawings drawings, XmlNode node, string topPath, string nvPrPath, ExcelGroupShape parent=null) :
        base(drawings, node, topPath, nvPrPath, parent)
    {
        this.Init(string.IsNullOrEmpty(this._topPath) ? "" : this._topPath + "/");
    }
    private void Init(string topPath)
    {
        this._shapeStylePath = string.Format(this._shapeStylePath, topPath);
        this._fillPath = string.Format(this._fillPath, topPath);
        this._borderPath = string.Format(this._borderPath, topPath);
        this._effectPath = string.Format(this._effectPath, topPath);
        this._headEndPath = string.Format(this._headEndPath, topPath);
        this._tailEndPath = string.Format(this._tailEndPath, topPath);
        this._textPath = string.Format(this._textPath, topPath);
        this._lockTextPath = string.Format(this._lockTextPath, topPath);
        this._textAnchoringPath = string.Format(this._textAnchoringPath, topPath);
        this._textAnchoringCtlPath = string.Format(this._textAnchoringCtlPath, topPath);
        this._paragraphPath = string.Format(this._paragraphPath, topPath);
        this._textAlignPath = string.Format(this._textAlignPath, topPath);
        this._indentAlignPath = string.Format(this._indentAlignPath, topPath);
        this._textVerticalPath = string.Format(this._textVerticalPath, topPath);
        this._fontPath = string.Format(this._fontPath, topPath);
        this._textBodyPath = string.Format(this._textBodyPath, topPath);
        this.AddSchemaNodeOrder(this.SchemaNodeOrder, new string[] { "nvSpPr", "spPr", "txSp", "style", "txBody", "hlinkClick", "hlinkHover", "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "blipFill", "gradFill", "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "pPr", "r", "br", "fld", "endParaRPr", "lnRef", "fillRef", "effectRef", "fontRef" });
    }
    /// <summary>
    /// The type of drawing
    /// </summary>
    public override eDrawingType DrawingType
    {
        get
        {
            return eDrawingType.Shape;
        }
    }
    /// <summary>
    /// Shape style
    /// </summary>
    public virtual eShapeStyle Style
    {
        get
        {
            string v = this.GetXmlNodeString(this._shapeStylePath);
            try
            {
                return (eShapeStyle)Enum.Parse(typeof(eShapeStyle), v, true);
            }
            catch
            {
                throw (new Exception(string.Format("Invalid shapetype {0}", v)));
            }
        }
        set
        {
            string v = value.ToString();
            v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
            this.SetXmlNodeString(this._shapeStylePath, v);
        }
    }
    ExcelDrawingFill _fill = null;
    /// <summary>
    /// Access Fill properties
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            return this._fill ??= new ExcelDrawingFill(this._drawings, this.NameSpaceManager, this.TopNode, this._fillPath, this.SchemaNodeOrder);
        }
    }
    ExcelDrawingBorder _border = null;
    /// <summary>
    /// Access to Border propesties
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get
        {
            return this._border ??= new ExcelDrawingBorder(this._drawings, this.NameSpaceManager, this.TopNode, this._borderPath, this.SchemaNodeOrder);
        }
    }
    ExcelDrawingEffectStyle _effect = null;
    /// <summary>
    /// Drawing effect properties
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            return this._effect ??= new ExcelDrawingEffectStyle(this._drawings,
                                                                this.NameSpaceManager,
                                                                this.TopNode,
                                                                this._effectPath,
                                                                this.SchemaNodeOrder);
        }
    }
    ExcelDrawing3D _threeD = null;
    /// <summary>
    /// Defines 3D properties to apply to an object
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get { return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, this._fillPath, this.SchemaNodeOrder); }
    }
    ExcelDrawingLineEnd _headEnd = null;
    /// <summary>
    /// Head line end
    /// </summary>
    public ExcelDrawingLineEnd HeadEnd
    {
        get { return this._headEnd ??= new ExcelDrawingLineEnd(this.NameSpaceManager, this.TopNode, this._headEndPath, this.InitSpPr); }
    }
    ExcelDrawingLineEnd _tailEnd = null;
    /// <summary>
    /// Tail line end
    /// </summary>
    public ExcelDrawingLineEnd TailEnd
    {
        get { return this._tailEnd ??= new ExcelDrawingLineEnd(this.NameSpaceManager, this.TopNode, this._tailEndPath, this.InitSpPr); }
    }
    ExcelTextFont _font = null;
    /// <summary>
    /// Font properties
    /// </summary>
    public ExcelTextFont Font
    {
        get
        {
            if (this._font == null)
            {
                XmlNode node = this.TopNode.SelectSingleNode(this._paragraphPath, this.NameSpaceManager);
                if (node == null)
                {
                    this.Text = "";    //Creates the node p element
                    node = this.TopNode.SelectSingleNode(this._paragraphPath, this.NameSpaceManager);
                }

                this._font = new ExcelTextFont(this._drawings, this.NameSpaceManager, this.TopNode, this._fontPath, this.SchemaNodeOrder);
            }
            return this._font;
        }
    }
    bool isSpInit = false;
    private void InitSpPr()
    {
        if (this.isSpInit == false)
        {
            if (this.CreateNodeUntil(this._topPath, "spPr", out XmlNode spPrNode))
            {
                spPrNode.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln ><a:effectLst/><a:sp3d/>";
            }
        }

        this.isSpInit = true;
    }


    /// <summary>
    /// Text inside the shape
    /// </summary>
    public string Text
    {
        get
        {
            return this.RichText.Text;
        }
        set
        {
            this.RichText.Text = value;
        }

    }
    /// <summary>
    /// Lock drawing
    /// </summary>
    public bool LockText
    {
        get
        {
            return this.GetXmlNodeBool(this._lockTextPath, true);
        }
        set
        {
            this.SetXmlNodeBool(this._lockTextPath, value);
        }
    }
    ExcelParagraphCollection _richText = null;
    internal static string[] _shapeNodeOrder= new string[] { "ln", "headEnd", "tailEnd", "effectLst", "blur", "fillOverlay", "glow", "innerShdw", "outerShdw", "prstShdw", "reflection", "softEdges", "effectDag", "scene3d", "scene3D", "sp3d", "bevelT", "bevelB", "extrusionClr", "contourClr" };

    /// <summary>
    /// Richtext collection. Used to format specific parts of the text
    /// </summary>
    public ExcelParagraphCollection RichText
    {
        get
        {
            return this._richText ??= new ExcelParagraphCollection(this, this.NameSpaceManager, this.TopNode, this._paragraphPath, this.SchemaNodeOrder);
        }
    }
    /// <summary>
    /// Text Anchoring
    /// </summary>
    public eTextAnchoringType TextAnchoring
    {
        get
        {
            return this.GetXmlNodeString(this._textAnchoringPath).TranslateTextAchoring();
        }
        set
        {
            this.SetXmlNodeString(this._textAnchoringPath, value.TranslateTextAchoringText());
        }
    }
    /// <summary>
    /// The centering of the text box.
    /// </summary>
    public bool TextAnchoringControl
    {
        get
        {
            return this.GetXmlNodeBool(this._textAnchoringCtlPath);
        }
        set
        {
            if (value)
            {
                this.SetXmlNodeString(this._textAnchoringCtlPath, "1");
            }
            else
            {
                this.SetXmlNodeString(this._textAnchoringCtlPath, "0");
            }
        }
    }
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
    /// <summary>
    /// Indentation
    /// </summary>
    public int Indent
    {
        get
        {
            return this.GetXmlNodeInt(this._indentAlignPath);
        }
        set
        {
            if (value < 0 || value > 8)
            {
                throw (new ArgumentOutOfRangeException("Indent level must be between 0 and 8"));
            }

            this.SetXmlNodeString(this._indentAlignPath, value.ToString());
        }
    }
    /// <summary>
    /// Vertical text
    /// </summary>
    public eTextVerticalType TextVertical
    {
        get
        {
            return this.GetXmlNodeString(this._textVerticalPath).TranslateTextVertical();
        }
        set
        {
            this.SetXmlNodeString(this._textVerticalPath, value.TranslateTextVerticalText());
        }
    }
    ExcelTextBody _textBody = null;
    /// <summary>
    /// Access to text body properties.
    /// </summary>
    public ExcelTextBody TextBody
    {
        get
        {
            return this._textBody ??= new ExcelTextBody(this.NameSpaceManager, this.TopNode, this._textBodyPath, this.SchemaNodeOrder);
        }
    }

    internal override void CellAnchorChanged()
    {
        base.CellAnchorChanged();
        if (this._fill != null)
        {
            this._fill.SetTopNode(this.TopNode);
        }

        if (this._border != null)
        {
            this._border.TopNode = this.TopNode;
        }

        if (this._effect != null)
        {
            this._effect.TopNode = this.TopNode;
        }

        if (this._font != null)
        {
            this._font.TopNode = this.TopNode;
        }

        if (this._threeD != null)
        {
            this._threeD.TopNode = this.TopNode;
        }

        if (this._tailEnd != null)
        {
            this._tailEnd.TopNode = this.TopNode;
        }

        if (this._headEnd != null)
        {
            this._headEnd.TopNode = this.TopNode;
        }

        if (this._richText != null)
        {
            this._richText.TopNode = this.TopNode;
        }

        if (this._textBody != null)
        {
            this._textBody.TopNode = this.TopNode;
        }
    }
}