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
using OfficeOpenXml.Drawing;
using System.Drawing;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Used by Rich-text and Paragraphs.
    /// </summary>
    public class ExcelTextFont : XmlHelper
    {
        string _path;
        internal XmlNode _rootNode;
        Action _initXml;
        IPictureRelationDocument _pictureRelationDocument;
        internal ExcelTextFont(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder, Action initXml=null)
            : base(namespaceManager, rootNode)
        {
            this.AddSchemaNodeOrder(schemaNodeOrder, new string[] { "bodyPr", "lstStyle","p", "pPr", "defRPr", "solidFill","highlight", "uFill", "latin","ea", "cs","sym","hlinkClick","hlinkMouseOver","rtl", "r", "rPr", "t" });
            this._rootNode = rootNode;
            this._pictureRelationDocument = pictureRelationDocument;
            this._initXml = initXml;
            if (path != "")
            {
                XmlNode node = rootNode.SelectSingleNode(path, namespaceManager);
                if (node != null)
                {
                    this.TopNode = node;
                }
            }

            this._path = path;
        }
        string _fontLatinPath = "a:latin/@typeface";
        /// <summary>
        /// The latin typeface name
        /// </summary>
        public string LatinFont
        {
            get
            {
                return this.GetXmlNodeString(this._fontLatinPath);
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._fontLatinPath, value);
            }
        }
        string _fontEaPath = "a:ea/@typeface";
        /// <summary>
        /// The East Asian typeface name
        /// </summary>
        public string EastAsianFont
        {
            get
            {
                return this.GetXmlNodeString(this._fontEaPath);
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._fontEaPath, value);
            }
        }
        string _fontCsPath = "a:cs/@typeface";
        /// <summary>
        /// The complex font typeface name
        /// </summary>
        public string ComplexFont
        {
            get
            {
                return this.GetXmlNodeString(this._fontCsPath);
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._fontCsPath, value);
            }
        }

        /// <summary>
        /// Creates the top nodes of the collection
        /// </summary>
        protected internal void CreateTopNode()
        {
            if (this._path!="" && this.TopNode== this._rootNode)
            {
                this._initXml?.Invoke();
                this.CreateNode(this._path);
                this.TopNode = this._rootNode.SelectSingleNode(this._path, this.NameSpaceManager);
                this.CreateNode("../../../a:bodyPr");
                this.CreateNode("../../../a:lstStyle");
            }
            else if (this.TopNode.ParentNode?.ParentNode?.ParentNode?.LocalName == "rich")
            {
                this.CreateNode("../../../a:bodyPr");
                this.CreateNode("../../../a:lstStyle");
            }
        }
        string _boldPath = "@b";
        /// <summary>
        /// If the font is bold
        /// </summary>
        public bool Bold
        {
            get
            {
                return this.GetXmlNodeBool(this._boldPath);
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._boldPath, value ? "1" : "0");
            }
        }
        string _underLinePath = "@u";
        /// <summary>
        /// The fonts underline style
        /// </summary>
        public eUnderLineType UnderLine
        {
            get
            {
                return this.GetXmlNodeString(this._underLinePath).TranslateUnderline();
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._underLinePath, value.TranslateUnderlineText());
            }
        }

        internal void SetFromXml(XmlElement copyFromElement)
        {
            this.CreateTopNode();
            foreach (XmlAttribute a in copyFromElement.Attributes)
            {
                ((XmlElement)this.TopNode).SetAttribute(a.Name, a.NamespaceURI, a.Value);
            }
            if(copyFromElement.HasChildNodes && !this.TopNode.HasChildNodes)
            {
                this.TopNode.InnerXml = copyFromElement.InnerXml;
            }
        }

        string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// The fonts underline color
        /// </summary>
        public Color UnderLineColor
        {
            get
            {
                string col = this.GetXmlNodeString(this._underLineColorPath);
                if (col == "")
                {
                    return Color.Empty;
                }
                else
                {
                    return Color.FromArgb(int.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier));
                }
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._underLineColorPath, value.ToArgb().ToString("X").Substring(2, 6));
            }
        }
        string _italicPath = "@i";
        /// <summary>
        /// If the font is italic
        /// </summary>
        public bool Italic
        {
            get
            {
                return this.GetXmlNodeBool(this._italicPath);
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._italicPath, value ? "1" : "0");
            }
        }
        string _strikePath = "@strike";
        /// <summary>
        /// Font strike out type
        /// </summary>
        public eStrikeType Strike
        {
            get
            {
                return this.GetXmlNodeString(this._strikePath).TranslateStrikeType();
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._strikePath, value.TranslateStrikeTypeText());
            }
        }
        string _sizePath = "@sz";
        /// <summary>
        /// Font size
        /// </summary>
        public float Size
        {
            get
            {
                return this.GetXmlNodeInt(this._sizePath) / 100;
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeString(this._sizePath, ((int)(value * 100)).ToString());
            }
        }
        ExcelDrawingFill _fill;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                return this._fill ??= new ExcelDrawingFill(this._pictureRelationDocument,
                                                           this.NameSpaceManager,
                                                           this._rootNode,
                                                           this._path,
                                                           this.SchemaNodeOrder,
                                                           this.CreateTopNode);
            }
        }
        string _colorPath = "a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Sets the default color of the text.
        /// This sets the Fill to a SolidFill with the specified color.
        /// <remark>
        /// Use the Fill property for more options
        /// </remark>
        /// </summary>
        [Obsolete("Use the Fill property for more options")]
        public Color Color
        {
            get
            {
                string col = this.GetXmlNodeString(this._colorPath);
                if (col == "")
                {
                    return Color.Empty;
                }
                else
                {
                    return Color.FromArgb(int.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier));
                }
            }
            set
            {
                this.Fill.Style = eFillStyle.SolidFill;
                this.Fill.SolidFill.Color.SetRgbColor(value);
            }
        }
        string _kernPath = "@kern";
        /// <summary>
        /// Specifies the minimum font size at which character kerning occurs for this text run
        /// </summary>
        public double Kerning
        {
            get
            {
                return this.GetXmlNodeFontSize(this._kernPath);
            }
            set
            {
                this.CreateTopNode();
                this.SetXmlNodeFontSize(this._kernPath, value, "Kerning");
            }
        }

        /// <summary>
        /// Set the font style properties
        /// </summary>
        /// <param name="name">Font family name</param>
        /// <param name="size">Font size</param>
        /// <param name="bold"></param>
        /// <param name="italic"></param>
        /// <param name="underline"></param>
        /// <param name="strikeout"></param>
        public void SetFromFont(string name, float size, bool bold = false, bool italic = false, bool underline = false, bool strikeout = false)
        {
            this.LatinFont = name;
            this.ComplexFont = name;
            this.Size = size;
            if (bold)
            {
                this.Bold = bold;
            }

            if (italic)
            {
                this.Italic = italic;
            }

            if (underline)
            {
                this.UnderLine = eUnderLineType.Single;
            }

            if (strikeout)
            {
                this.Strike = eStrikeType.Single;
            }
        }
    }
}
