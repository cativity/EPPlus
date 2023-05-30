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
using System.Drawing;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Fill properties for drawing objects like lines etc, that don't have blip- and pattern- fills
/// </summary>
public class ExcelDrawingFillBasic : XmlHelper, IDisposable
{
    /// <summary>
    /// XPath
    /// </summary>
    internal protected string _fillPath;

    /// <summary>
    /// The fill xml element
    /// </summary>
    internal protected XmlNode _fillNode;

    /// <summary>
    /// The drawings collection
    /// </summary>
    internal protected ExcelDrawing _drawing;

    /// <summary>
    /// The fill type node.
    /// </summary>
    internal protected XmlNode _fillTypeNode;

    internal Action _initXml;

    internal ExcelDrawingFillBasic(ExcelPackage pck,
                                   XmlNamespaceManager nameSpaceManager,
                                   XmlNode topNode,
                                   string fillPath,
                                   string[] schemaNodeOrderBefore,
                                   bool doLoad,
                                   Action initXml = null)
        : base(nameSpaceManager, topNode)
    {
        this.AddSchemaNodeOrder(schemaNodeOrderBefore,
                                new string[]
                                {
                                    "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill", "grpFill", "ln",
                                    "effectLst", "effectDag", "highlight", "latin", "cs", "sym", "ea", "hlinkClick", "hlinkMouseOver", "rtl"
                                });

        this._fillPath = fillPath;
        this._initXml = initXml;
        this.SetFillNodes(topNode);

        //Setfill node
        if (doLoad && this._fillNode != null)
        {
            this.LoadFill();
        }

        if (pck != null)
        {
            pck.BeforeSave.Add(this.BeforeSave);
        }
    }

    internal void SetTopNode(XmlNode topNode)
    {
        this.TopNode = topNode;
        this.SetFillNodes(topNode);
        this._fillTypeNode = null;
        this.LoadFill();
    }

    private void SetFillNodes(XmlNode topNode)
    {
        if (string.IsNullOrEmpty(this._fillPath))
        {
            this._fillNode = topNode;

            if (topNode.LocalName.EndsWith("Fill")) //Theme nodes will have the fillnode as topnode
            {
                this._fillTypeNode = this._fillNode;
            }
        }
        else
        {
            this._fillNode = topNode.SelectSingleNode(this._fillPath, this.NameSpaceManager);
        }
    }

    internal virtual void BeforeSave()
    {
        if (this._gradientFill != null)
        {
            this._gradientFill.UpdateXml();
        }
    }

    /// <summary>
    /// Loads the fill from xml
    /// </summary>
    internal protected virtual void LoadFill()
    {
        this._fillTypeNode ??=
            (this._fillNode.SelectSingleNode("a:solidFill", this.NameSpaceManager) ?? this._fillNode.SelectSingleNode("a:gradFill", this.NameSpaceManager))
            ?? this._fillNode.SelectSingleNode("a:noFill", this.NameSpaceManager);

        if (this._fillTypeNode == null)
        {
            return;
        }

        switch (this._fillTypeNode.LocalName)
        {
            case "solidFill":
                this._style = eFillStyle.SolidFill;
                this._solidFill = new ExcelDrawingSolidFill(this.NameSpaceManager, this._fillTypeNode, "", this.SchemaNodeOrder, this._initXml);

                break;

            case "gradFill":
                this._style = eFillStyle.GradientFill;
                this._gradientFill = new ExcelDrawingGradientFill(this.NameSpaceManager, this._fillTypeNode, this.SchemaNodeOrder, this._initXml);

                break;

            default:
                this._style = eFillStyle.NoFill;

                break;
        }
    }

    internal void SetFromXml(ExcelDrawingFill fill)
    {
        this.Style = fill.Style;
        XmlElement? copyFromFillElement = (XmlElement)fill._fillTypeNode;

        foreach (XmlAttribute a in copyFromFillElement.Attributes)
        {
            _ = ((XmlElement)this._fillTypeNode).SetAttribute(a.Name, a.NamespaceURI, a.Value);
        }

        this._fillTypeNode.InnerXml = copyFromFillElement.InnerXml;

        if (fill.Style == eFillStyle.BlipFill)
        {
            XmlAttribute relAttr = (XmlAttribute)this._fillTypeNode.SelectSingleNode("a:blip/@r:embed", this.NameSpaceManager);

            if (relAttr?.Value != null)
            {
                _ = relAttr.OwnerElement.Attributes.Remove(relAttr);
            }
        }

        this.LoadFill();

        if (this.Style == eFillStyle.BlipFill && fill.BlipFill.Image.ImageBytes != null)
        {
            _ = ((ExcelDrawingFill)this).BlipFill.Image.SetImage(fill.BlipFill.Image.ImageBytes, fill.BlipFill.Image.Type ?? ePictureType.Jpg);
        }
    }

    private static void CreateImageRelation(ExcelDrawingFill fill, XmlElement copyFromFillElement)
    {
    }

    internal string GetFromXml()
    {
        return this._fillTypeNode.OuterXml;
    }

    internal virtual void SetFillProperty()
    {
        if (this._fillNode == null)
        {
            this.InitSpPr(eFillStyle.SolidFill);
            this.Style = eFillStyle.SolidFill; //This will create the _fillNode
            this._solidFill = new ExcelDrawingSolidFill(this.NameSpaceManager, this._fillTypeNode, "", this.SchemaNodeOrder, this._initXml);

            return;
        }

        this._solidFill = null;
        this._gradientFill = null;

        switch (this._fillTypeNode.LocalName)
        {
            case "solidFill":
                this._solidFill = new ExcelDrawingSolidFill(this.NameSpaceManager, this._fillTypeNode, "", this.SchemaNodeOrder, this._initXml);

                break;

            case "gradFill":
                this._gradientFill = new ExcelDrawingGradientFill(this.NameSpaceManager, this._fillTypeNode, this.SchemaNodeOrder, this._initXml);

                break;

            default:
                if (this is ExcelDrawingFillBasic && this._style != eFillStyle.NoFill)
                {
                    throw new ArgumentException("Style", $"Style {this.Style} cannot be applied to this object.");
                }

                break;
        }
    }

    bool isSpInit;

    private void InitSpPr(eFillStyle style)
    {
        if (this.isSpInit == false)
        {
            if (!string.IsNullOrEmpty(this._fillPath) && !this.ExistsNode(this._fillPath) && this.CreateNodeUntil(this._fillPath, "spPr", out XmlNode spPrNode))
            {
                if (this._fillPath.EndsWith("ln"))
                {
                    spPrNode.InnerXml = $"<a:ln><a:noFill/></a:ln ><a:effectLst/><a:sp3d/>";
                }
                else
                {
                    spPrNode.InnerXml = $"<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/>";
                }

                this._fillNode = this.GetNode(this._fillPath);
                this._fillTypeNode = this._fillNode.FirstChild;
            }
            else if (this._fillTypeNode == null)
            {
                this._fillNode ??= this.GetNode(this._fillPath);

                if (!this._fillNode.HasChildNodes)
                {
                    this._fillNode.InnerXml = $"<a:{GetStyleText(style)}/>";
                }

                this.LoadFill();
            }
        }

        this.isSpInit = true;
    }

    internal eFillStyle? _style;

    /// <summary>
    /// Fill style
    /// </summary>
    public eFillStyle Style
    {
        get { return this._style ?? eFillStyle.NoFill; }
        set
        {
            if (this._style == value)
            {
                return;
            }

            this._initXml?.Invoke();

            if (value == eFillStyle.GroupFill)
            {
                throw new NotImplementedException("Fillstyle not implemented");
            }
            else
            {
                this._style = value;
                this.InitSpPr(value);
                this.CreateFillTopNode(value);
                this.SetFillProperty();
            }
        }
    }

    const string ColorPath = "a:srgbClr/@val";

    /// <summary>
    /// Fill color for solid fills.
    /// Other fill styles will return Color.Empty.
    /// Setting this propery will set the Type to SolidFill with the specified color.
    /// </summary>
    public Color Color
    {
        get
        {
            if (this.Style != eFillStyle.SolidFill)
            {
                return Color.Empty;
            }

            if (this.SolidFill.Color.ColorType != eDrawingColorType.Rgb)
            {
                return Color.Empty;
            }

            Color col = this.SolidFill.Color.RgbColor.Color;

            if (col == Color.Empty)
            {
                return Color.FromArgb(79, 129, 189);
            }
            else
            {
                return col;
            }
        }
        set
        {
            this._initXml?.Invoke();
            this.Style = eFillStyle.SolidFill;
            this.SolidFill.Color.SetRgbColor(value);
        }
    }

    private ExcelDrawingSolidFill _solidFill;

    /// <summary>
    /// Reference solid fill properties
    /// This property is only accessable when Type is set to SolidFill
    /// </summary>
    public ExcelDrawingSolidFill SolidFill
    {
        get
        {
            if (this.Style == eFillStyle.SolidFill && this._solidFill == null)
            {
                this._solidFill = new ExcelDrawingSolidFill(this.NameSpaceManager, this._fillTypeNode, "", this.SchemaNodeOrder, this._initXml);
            }

            return this._solidFill;
        }
    }

    private ExcelDrawingGradientFill _gradientFill;

    /// <summary>
    /// Reference gradient fill properties
    /// This property is only accessable when Type is set to GradientFill
    /// </summary>
    public ExcelDrawingGradientFill GradientFill
    {
        get { return this._gradientFill; }
    }

    /// <summary>
    /// Transparancy in percent from a solid fill. 
    /// This is the same as 100-Fill.Transform.Alpha
    /// </summary>
    public int Transparancy
    {
        get
        {
            if (this._solidFill == null)
            {
                return 0;
            }

            return (int)(100 - this._solidFill.Color.Transforms.FindValue(eColorTransformType.Alpha));
        }
        set
        {
            if (this._solidFill == null)
            {
                throw new InvalidOperationException("Transparency can only be set when Type is set to SolidFill.");
            }

            IColorTransformItem? alphaItem = this._solidFill.Color.Transforms.Find(eColorTransformType.Alpha);

            if (alphaItem == null)
            {
                this._solidFill.Color.Transforms.AddAlpha(100 - value);
            }
            else
            {
                alphaItem.Value = 100 - value;
            }
        }
    }

    private void CreateFillTopNode(eFillStyle value)
    {
        if (this._fillNode == this.TopNode)
        {
            if (this._fillNode == this._fillTypeNode)
            {
                XmlElement? node = this._fillTypeNode.OwnerDocument.CreateElement("a", GetStyleText(value), ExcelPackage.schemaDrawings);
                _ = this._fillTypeNode.ParentNode.InsertBefore(node, this._fillTypeNode);
                _ = this._fillTypeNode.ParentNode.RemoveChild(this._fillTypeNode);
                this._fillTypeNode = node;
                this._fillNode = node;
                this.TopNode = node;
            }
            else
            {
                this._fillTypeNode = this.CreateNode("a:" + GetStyleText(value));
            }
        }
        else
        {
            if (this._fillTypeNode != null)
            {
                _ = this._fillTypeNode.ParentNode.RemoveChild(this._fillTypeNode);
            }

            this._fillTypeNode = this.CreateNode(this._fillPath + "/a:" + GetStyleText(value), false);
            this._fillNode ??= this._fillTypeNode.ParentNode;
        }
    }

    internal static eFillStyle GetStyleEnum(string name)
    {
        switch (name)
        {
            case "noFill":
                return eFillStyle.NoFill;

            case "blipFill":
                return eFillStyle.BlipFill;

            case "gradFill":
                return eFillStyle.GradientFill;

            case "grpFill":
                return eFillStyle.GroupFill;

            case "pattFill":
                return eFillStyle.PatternFill;

            default:
                return eFillStyle.SolidFill;
        }
    }

    internal static string GetStyleText(eFillStyle style)
    {
        switch (style)
        {
            case eFillStyle.BlipFill:
                return "blipFill";

            case eFillStyle.GradientFill:
                return "gradFill";

            case eFillStyle.GroupFill:
                return "grpFill";

            case eFillStyle.NoFill:
                return "noFill";

            case eFillStyle.PatternFill:
                return "pattFill";

            default:
                return "solidFill";
        }
    }

    /// <summary>
    /// Disposes the object
    /// </summary>
    public void Dispose()
    {
        this._fillNode = null;
        this._solidFill = null;
        this._gradientFill = null;
    }

    internal void UpdateFillTypeNode()
    {
        if (this._fillTypeNode != null && this._fillTypeNode.ParentNode == null)
        {
            this._fillTypeNode = null;
            this.LoadFill();
        }
    }
}