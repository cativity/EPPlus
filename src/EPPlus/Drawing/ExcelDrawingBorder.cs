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

using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Border for drawings
/// </summary>    
public class ExcelDrawingBorder : XmlHelper
{
    string _linePath;
    IPictureRelationDocument _pictureRelationDocument;
    bool isSpInit;

    internal ExcelDrawingBorder(IPictureRelationDocument pictureRelationDocument,
                                XmlNamespaceManager nameSpaceManager,
                                XmlNode topNode,
                                string linePath,
                                string[] schemaNodeOrder)
        : base(nameSpaceManager, topNode)
    {
        this.AddSchemaNodeOrder(schemaNodeOrder,
                                new string[]
                                {
                                    "noFill", "solidFill", "gradFill", "pattFill", "prstDash", "custDash", "round", "bevel", "miter", "headEnd", "tailEnd"
                                });

        this._linePath = linePath;
        this._lineStylePath = string.Format(this._lineStylePath, linePath);
        this._lineCapPath = string.Format(this._lineCapPath, linePath);
        this._lineWidth = string.Format(this._lineWidth, linePath);
        this._bevelPath = string.Format(this._bevelPath, linePath);
        this._roundPath = string.Format(this._roundPath, linePath);
        this._miterPath = string.Format(this._miterPath, linePath);
        this._miterJoinLimitPath = string.Format(this._miterJoinLimitPath, linePath);

        this._headEndPath = string.Format(this._headEndPath, linePath);
        this._tailEndPath = string.Format(this._tailEndPath, linePath);
        this._compoundLineTypePath = string.Format(this._compoundLineTypePath, linePath);
        this._alignmentPath = string.Format(this._alignmentPath, linePath);
        this._pictureRelationDocument = pictureRelationDocument;
    }

    #region "Public properties"

    ExcelDrawingFillBasic _fill;

    /// <summary>
    /// Access to fill properties
    /// </summary>
    public ExcelDrawingFillBasic Fill
    {
        get
        {
            return this._fill ??= new ExcelDrawingFillBasic(this._pictureRelationDocument.Package,
                                                            this.NameSpaceManager,
                                                            this.TopNode,
                                                            this._linePath,
                                                            this.SchemaNodeOrder,
                                                            true);
        }
    }

    string _lineStylePath = "{0}/a:prstDash/@val";

    /// <summary>
    /// Preset line dash
    /// </summary>
    public eLineStyle? LineStyle
    {
        get
        {
            string? v = this.GetXmlNodeString(this._lineStylePath);

            if (string.IsNullOrEmpty(v))
            {
                return null;
            }
            else
            {
                return EnumTransl.ToLineStyle(v);
            }
        }
        set
        {
            this.InitSpPr();
            _ = this.CreateNode(this._linePath, false);

            if (value.HasValue)
            {
                this.SetXmlNodeString(this._lineStylePath, EnumTransl.FromLineStyle(value.Value));
            }
            else
            {
                this.DeleteNode(this._lineStylePath, true);
            }
        }
    }

    private void InitSpPr()
    {
        if (this.isSpInit == false)
        {
            if (this.CreateNodeUntil(this._linePath, "spPr", out XmlNode spPrNode))
            {
                spPrNode.InnerXml = "<a:ln><a:noFill/></a:ln ><a:effectLst/>";
            }
        }

        this.isSpInit = true;
    }

    string _compoundLineTypePath = "{0}/@cmpd";

    /// <summary>
    /// The compound line type that is to be used for lines with text such as underlines
    /// </summary>
    public eCompundLineStyle CompoundLineStyle
    {
        get { return EnumTransl.ToLineCompound(this.GetXmlNodeString(this._compoundLineTypePath)); }
        set
        {
            this.InitSpPr();
            this.SetXmlNodeString(this._compoundLineTypePath, EnumTransl.FromLineCompound(value));
        }
    }

    string _alignmentPath = "{0}/@algn";

    /// <summary>
    /// The pen alignment type for use within a text body
    /// </summary>
    public ePenAlignment Alignment
    {
        get { return EnumTransl.ToPenAlignment(this.GetXmlNodeString(this._alignmentPath)); }
        set
        {
            this.InitSpPr();
            this.SetXmlNodeString(this._alignmentPath, EnumTransl.FromPenAlignment(value));
        }
    }

    string _lineCapPath = "{0}/@cap";

    /// <summary>
    /// Specifies how to cap the ends of lines
    /// </summary>
    public eLineCap LineCap
    {
        get { return EnumTransl.ToLineCap(this.GetXmlNodeString(this._lineCapPath)); }
        set
        {
            this.InitSpPr();
            this.SetXmlNodeString(this._lineCapPath, EnumTransl.FromLineCap(value));
        }
    }

    string _lineWidth = "{0}/@w";

    /// <summary>
    /// Width in pixels
    /// </summary>
    public double Width
    {
        get { return this.GetXmlNodeEmuToPt(this._lineWidth); }
        set
        {
            this.InitSpPr();
            this.SetXmlNodeEmuToPt(this._lineWidth, value);
        }
    }

    string _bevelPath = "{0}/a:bevel";
    string _roundPath = "{0}/a:round";
    string _miterPath = "{0}/a:miter";

    /// <summary>
    /// How connected lines are joined
    /// </summary>
    public eLineJoin? Join
    {
        get
        {
            if (this.ExistsNode(this._bevelPath))
            {
                return eLineJoin.Bevel;
            }
            else if (this.ExistsNode(this._roundPath))
            {
                return eLineJoin.Round;
            }
            else if (this.ExistsNode(this._miterPath))
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
            this.InitSpPr();

            if (value == eLineJoin.Bevel)
            {
                _ = this.CreateNode(this._bevelPath);
                this.DeleteNode(this._roundPath);
                this.DeleteNode(this._miterPath);
            }
            else if (value == eLineJoin.Round)
            {
                _ = this.CreateNode(this._roundPath);
                this.DeleteNode(this._bevelPath);
                this.DeleteNode(this._miterPath);
            }
            else
            {
                _ = this.CreateNode(this._miterPath);
                this.DeleteNode(this._roundPath);
                this.DeleteNode(this._bevelPath);
            }
        }
    }

    string _miterJoinLimitPath = "{0}/a:miter/@lim";

    /// <summary>
    /// The amount by which lines is extended to form a miter join 
    /// Otherwise miter joins can extend infinitely far.
    /// </summary>
    public double? MiterJoinLimit
    {
        get { return this.GetXmlNodePercentage(this._miterJoinLimitPath); }
        set
        {
            this.Join = eLineJoin.Miter;
            this.SetXmlNodePercentage(this._miterJoinLimitPath, value, false, double.MaxValue);
        }
    }

    string _headEndPath = "{0}/a:headEnd";
    ExcelDrawingLineEnd _headEnd = null;

    /// <summary>
    /// Head end style for the line
    /// </summary>
    public ExcelDrawingLineEnd HeadEnd
    {
        get
        {
            if (this._headEnd == null)
            {
                return new ExcelDrawingLineEnd(this.NameSpaceManager, this.TopNode, this._headEndPath, this.InitSpPr);
            }

            return this._headEnd;
        }
    }

    string _tailEndPath = "{0}/a:tailEnd";
    ExcelDrawingLineEnd _tailEnd = null;

    /// <summary>
    /// Tail end style for the line
    /// </summary>
    public ExcelDrawingLineEnd TailEnd
    {
        get
        {
            if (this._tailEnd == null)
            {
                return new ExcelDrawingLineEnd(this.NameSpaceManager, this.TopNode, this._tailEndPath, this.InitSpPr);
            }

            return this._tailEnd;
        }
    }

    #endregion

    internal XmlElement LineElement
    {
        get { return this.TopNode.SelectSingleNode(this._linePath, this.NameSpaceManager) as XmlElement; }
    }

    internal void SetFromXml(XmlElement copyFromLineElement)
    {
        this.InitSpPr();
        XmlElement lineElement = this.LineElement;

        if (lineElement == null)
        {
            _ = this.CreateNode(this._linePath);
        }

        foreach (XmlAttribute a in copyFromLineElement.Attributes)
        {
            _ = lineElement.SetAttribute(a.Name, a.NamespaceURI, a.Value);
        }

        lineElement.InnerXml = copyFromLineElement.InnerXml;
    }
}