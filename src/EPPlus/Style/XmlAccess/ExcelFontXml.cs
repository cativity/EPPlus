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
using System.Collections.Generic;
#if NETFULL
using System.Drawing;
#endif
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

/// <summary>
/// Xml access class for fonts
/// </summary>
public sealed class ExcelFontXml : StyleXmlHelper
{
    internal ExcelFontXml(XmlNamespaceManager nameSpaceManager)
        : base(nameSpaceManager)
    {
        this._name = "";
        this._size = 0;
        this._family = int.MinValue;
        this._scheme = "";
        this._color = this._color = new ExcelColorXml(this.NameSpaceManager);
        this._bold = false;
        this._italic = false;
        this._strike = false;
        this._underlineType = ExcelUnderLineType.None;
        this._verticalAlign = "";
        this._charset = null;
    }

    internal ExcelFontXml(XmlNamespaceManager nsm, XmlNode topNode)
        : base(nsm, topNode)
    {
        this._name = this.GetXmlNodeString(namePath);
        this._size = (float)this.GetXmlNodeDecimal(sizePath);
        this._family = this.GetXmlNodeIntNull(familyPath) ?? int.MinValue;
        this._scheme = this.GetXmlNodeString(schemePath);
        this._color = new ExcelColorXml(nsm, topNode.SelectSingleNode(_colorPath, nsm));
        this._bold = this.GetBoolValue(topNode, boldPath);
        this._italic = this.GetBoolValue(topNode, italicPath);
        this._strike = this.GetBoolValue(topNode, strikePath);
        this._verticalAlign = this.GetXmlNodeString(verticalAlignPath);
        this._charset = this.GetXmlNodeIntNull(_charsetPath);

        if (topNode.SelectSingleNode(underLinedPath, this.NameSpaceManager) != null)
        {
            string ut = this.GetXmlNodeString(underLinedPath + "/@val");

            if (ut == "")
            {
                this._underlineType = ExcelUnderLineType.Single;
            }
            else
            {
                this._underlineType = (ExcelUnderLineType)Enum.Parse(typeof(ExcelUnderLineType), ut, true);
            }
        }
        else
        {
            this._underlineType = ExcelUnderLineType.None;
        }
    }

    internal override string Id
    {
        get
        {
            return this.Name
                   + "|"
                   + this.Size
                   + "|"
                   + this.Family
                   + "|"
                   + this.Color.Id
                   + "|"
                   + this.Scheme
                   + "|"
                   + this.Bold.ToString()
                   + "|"
                   + this.Italic.ToString()
                   + "|"
                   + this.Strike.ToString()
                   + "|"
                   + this.VerticalAlign
                   + "|"
                   + this.UnderLineType.ToString()
                   + "|"
                   + (this.Charset.HasValue ? this.Charset.ToString() : "");
        }
    }

    const string namePath = "d:name/@val";
    string _name;

    /// <summary>
    /// The name of the font
    /// </summary>
    public string Name
    {
        get { return this._name; }
        set
        {
            this.Scheme = ""; //Reset schema to avoid corrupt file if unsupported font is selected.
            this._name = value;
        }
    }

    const string sizePath = "d:sz/@val";
    float _size;

    /// <summary>
    /// Font size
    /// </summary>
    public float Size
    {
        get { return this._size; }
        set { this._size = value; }
    }

    const string familyPath = "d:family/@val";
    int _family;

    /// <summary>
    /// Font family
    /// </summary>
    public int Family
    {
        get { return this._family == int.MinValue ? 0 : this._family; }
        set { this._family = value; }
    }

    ExcelColorXml _color = null;
    const string _colorPath = "d:color";

    /// <summary>
    /// Text color
    /// </summary>
    public ExcelColorXml Color
    {
        get { return this._color; }
        internal set { this._color = value; }
    }

    const string schemePath = "d:scheme/@val";
    string _scheme = "";

    /// <summary>
    /// Font Scheme
    /// </summary>
    public string Scheme
    {
        get { return this._scheme; }
        internal set { this._scheme = value; }
    }

    const string boldPath = "d:b";
    bool _bold;

    /// <summary>
    /// If the font is bold
    /// </summary>
    public bool Bold
    {
        get { return this._bold; }
        set { this._bold = value; }
    }

    const string italicPath = "d:i";
    bool _italic;

    /// <summary>
    /// If the font is italic
    /// </summary>
    public bool Italic
    {
        get { return this._italic; }
        set { this._italic = value; }
    }

    const string strikePath = "d:strike";
    bool _strike;

    /// <summary>
    /// If the font is striked out
    /// </summary>
    public bool Strike
    {
        get { return this._strike; }
        set { this._strike = value; }
    }

    const string underLinedPath = "d:u";

    /// <summary>
    /// If the font is underlined.
    /// When set to true a the text is underlined with a single line
    /// </summary>
    public bool UnderLine
    {
        get { return this.UnderLineType != ExcelUnderLineType.None; }
        set { this._underlineType = value ? ExcelUnderLineType.Single : ExcelUnderLineType.None; }
    }

    ExcelUnderLineType _underlineType;

    /// <summary>
    /// If the font is underlined
    /// </summary>
    public ExcelUnderLineType UnderLineType
    {
        get { return this._underlineType; }
        set { this._underlineType = value; }
    }

    const string verticalAlignPath = "d:vertAlign/@val";
    string _verticalAlign;

    /// <summary>
    /// Vertical aligned
    /// </summary>
    public string VerticalAlign
    {
        get { return this._verticalAlign; }
        set { this._verticalAlign = value; }
    }

    const string _charsetPath = "d:charset/@val";
    int? _charset = null;

    /// <summary>
    /// The character set for the font
    /// </summary>
    /// <remarks>
    /// The following values can be used for this property.
    /// <list type="table">
    /// <listheader>Value</listheader><listheader>Description</listheader>
    /// <item>null</item><item>Not specified</item>
    /// <item>0x00</item><item>The ANSI character set. (IANA name iso-8859-1)</item>
    /// <item>0x01</item><item>The default character set.</item>
    /// <item>0x02</item><item>The Symbol character set. This value specifies that the characters in the Unicode private use area(U+FF00 to U+FFFF) of the font should be used to display characters in the range U+0000 to U+00FF.</item>       
    ///<item>0x4D</item><item>A Macintosh(Standard Roman) character set. (IANA name macintosh)</item>
    ///<item>0x80</item><item>The JIS character set. (IANA name shift_jis)</item>
    ///<item>0x81</item><item>The Hangul character set. (IANA name ks_c_5601-1987)</item>
    ///<item>0x82</item><item>A Johab character set. (IANA name KS C-5601-1992)</item>
    ///<item>0x86</item><item>The GB-2312 character set. (IANA name GBK)</item>
    ///<item>0x88</item><item>The Chinese Big Five character set. (IANA name Big5)</item>
    ///<item>0xA1</item><item>A Greek character set. (IANA name windows-1253)</item>
    ///<item>0xA2</item><item>A Turkish character set. (IANA name iso-8859-9)</item>
    ///<item>0xA3</item><item>A Vietnamese character set. (IANA name windows-1258)</item>
    ///<item>0xB1</item><item>A Hebrew character set. (IANA name windows-1255)</item>
    ///<item>0xB2</item><item>An Arabic character set. (IANA name windows-1256)</item>
    ///<item>0xBA</item><item>A Baltic character set. (IANA name windows-1257)</item>
    ///<item>0xCC</item><item>A Russian character set. (IANA name windows-1251)</item>
    ///<item>0xDE</item><item>A Thai character set. (IANA name windows-874)</item>
    ///<item>0xEE</item><item>An Eastern European character set. (IANA name windows-1250)</item>
    ///<item>0xFF</item><item>An OEM character set not defined by ISO/IEC 29500.</item>
    ///<item>Any other value</item><item>Application-defined, can be ignored</item>
    /// </list>
    /// </remarks>
    public int? Charset
    {
        get { return this._charset; }
        set { this._charset = value; }
    }

    /// <summary>
    /// Set the font properties
    /// </summary>
    /// <param name="name">Font family name</param>
    /// <param name="size">Font size</param>
    /// <param name="bold"></param>
    /// <param name="italic"></param>
    /// <param name="underline"></param>
    /// <param name="strikeout"></param>
    public void SetFromFont(string name, float size, bool bold = false, bool italic = false, bool underline = false, bool strikeout = false)
    {
        this.Name = name;

        //Family=fnt.FontFamily.;
        this.Size = size;
        this.Strike = strikeout;
        this.Bold = bold;
        this.UnderLine = underline;
        this.Italic = italic;
    }

    /// <summary>
    /// Gets the height of the font in 
    /// </summary>
    /// <param name="name"></param>
    /// <param name="size"></param>
    /// <returns></returns>
    internal static float GetFontHeight(string name, float size)
    {
        name = name.StartsWith("@") ? name.Substring(1) : name;

        return Convert.ToSingle(ExcelWorkbook.GetHeightPixels(name, size));
    }

    internal ExcelFontXml Copy()
    {
        ExcelFontXml newFont = new ExcelFontXml(this.NameSpaceManager);
        newFont.Name = this._name;
        newFont.Size = this._size;
        newFont.Family = this._family;
        newFont.Scheme = this._scheme;
        newFont.Bold = this._bold;
        newFont.Italic = this._italic;
        newFont.UnderLineType = this._underlineType;
        newFont.Strike = this._strike;
        newFont.VerticalAlign = this._verticalAlign;
        newFont.Color = this.Color.Copy();
        newFont.Charset = this._charset;

        return newFont;
    }

    internal override XmlNode CreateXmlNode(XmlNode topElement)
    {
        this.TopNode = topElement;

        if (this._bold)
        {
            _ = this.CreateNode(boldPath);
        }
        else
        {
            this.DeleteAllNode(boldPath);
        }

        if (this._italic)
        {
            _ = this.CreateNode(italicPath);
        }
        else
        {
            this.DeleteAllNode(italicPath);
        }

        if (this._strike)
        {
            _ = this.CreateNode(strikePath);
        }
        else
        {
            this.DeleteAllNode(strikePath);
        }

        if (this._underlineType == ExcelUnderLineType.None)
        {
            this.DeleteAllNode(underLinedPath);
        }
        else if (this._underlineType == ExcelUnderLineType.Single)
        {
            _ = this.CreateNode(underLinedPath);
        }
        else
        {
            string? v = this._underlineType.ToString();
            this.SetXmlNodeString(underLinedPath + "/@val", v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1));
        }

        if (this._verticalAlign != "")
        {
            this.SetXmlNodeString(verticalAlignPath, this._verticalAlign.ToString());
        }

        if (this._size > 0)
        {
            this.SetXmlNodeString(sizePath, this._size.ToString(CultureInfo.InvariantCulture));
        }

        if (this._color.Exists)
        {
            _ = this.CreateNode(_colorPath);
            _ = this.TopNode.AppendChild(this._color.CreateXmlNode(this.TopNode.SelectSingleNode(_colorPath, this.NameSpaceManager)));
        }

        if (!string.IsNullOrEmpty(this._name))
        {
            this.SetXmlNodeString(namePath, this._name);
        }

        if (this._family > int.MinValue)
        {
            this.SetXmlNodeString(familyPath, this._family.ToString());
        }

        this.SetXmlNodeInt(_charsetPath, this.Charset);

        if (this._scheme != "")
        {
            this.SetXmlNodeString(schemePath, this._scheme.ToString());
        }

        return this.TopNode;
    }
}