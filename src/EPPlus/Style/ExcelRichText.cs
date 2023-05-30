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
using System.Text;
using System.Xml;
using System.Drawing;
using System.Globalization;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;

namespace OfficeOpenXml.Style;

/// <summary>
/// A richtext part
/// </summary>
public class ExcelRichText : XmlHelper
{
    internal ExcelRichText(XmlNamespaceManager ns, XmlNode topNode, ExcelRichTextCollection collection)
        : base(ns, topNode)
    {
        this.SchemaNodeOrder = new string[] { "rPr", "t", "b", "i", "strike", "u", "vertAlign", "sz", "color", "rFont", "family", "scheme", "charset" };
        this._collection = collection;
    }

    internal delegate void CallbackDelegate();

    internal CallbackDelegate _callback;

    internal void SetCallback(CallbackDelegate callback)
    {
        this._callback = callback;
    }

    const string TEXT_PATH = "d:t";

    /// <summary>
    /// The text
    /// </summary>
    public string Text
    {
        get { return this.GetXmlNodeString(TEXT_PATH); }
        set
        {
            if (value == null)
            {
                throw new InvalidOperationException("Text can't be null");
            }

            this._collection.ConvertRichtext();
            this.SetXmlNodeString(TEXT_PATH, value, false);

            if (this.PreserveSpace)
            {
                XmlElement elem = this.TopNode.SelectSingleNode(TEXT_PATH, this.NameSpaceManager) as XmlElement;
                elem.SetAttribute("xml:space", "preserve");
            }

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    /// <summary>
    /// Returns the rich text item as a html string.
    /// </summary>
    public string HtmlText
    {
        get
        {
            StringBuilder? sb = new StringBuilder();
            this.WriteHtmlText(sb);

            return sb.ToString();
        }
    }

    internal void WriteHtmlText(StringBuilder sb)
    {
        _ = sb.Append("<span style=\"");
        HtmlRichText.GetRichTextStyle(this, sb);
        _ = sb.Append("\">");
        _ = sb.Append(this.Text);
        _ = sb.Append("</span>");
    }

    /// <summary>
    /// Preserves whitespace. Default true
    /// </summary>
    public bool PreserveSpace
    {
        get
        {
            XmlElement elem = this.TopNode.SelectSingleNode(TEXT_PATH, this.NameSpaceManager) as XmlElement;

            if (elem != null)
            {
                return elem.GetAttribute("xml:space") == "preserve";
            }

            return false;
        }
        set
        {
            this._collection.ConvertRichtext();
            XmlElement elem = this.TopNode.SelectSingleNode(TEXT_PATH, this.NameSpaceManager) as XmlElement;

            if (elem != null)
            {
                if (value)
                {
                    elem.SetAttribute("xml:space", "preserve");
                }
                else
                {
                    elem.RemoveAttribute("xml:space");
                }
            }

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    const string BOLD_PATH = "d:rPr/d:b";

    /// <summary>
    /// Bold text
    /// </summary>
    public bool Bold
    {
        get { return this.ExistsNode(BOLD_PATH); }
        set
        {
            this._collection.ConvertRichtext();

            if (value)
            {
                _ = this.CreateNode(BOLD_PATH);
            }
            else
            {
                this.DeleteNode(BOLD_PATH);
            }

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    const string ITALIC_PATH = "d:rPr/d:i";

    /// <summary>
    /// Italic text
    /// </summary>
    public bool Italic
    {
        get
        {
            //return GetXmlNodeBool(ITALIC_PATH, false);
            return this.ExistsNode(ITALIC_PATH);
        }
        set
        {
            this._collection.ConvertRichtext();

            if (value)
            {
                _ = this.CreateNode(ITALIC_PATH);
            }
            else
            {
                this.DeleteNode(ITALIC_PATH);
            }

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    const string STRIKE_PATH = "d:rPr/d:strike";

    /// <summary>
    /// Strike-out text
    /// </summary>
    public bool Strike
    {
        get { return this.ExistsNode(STRIKE_PATH); }
        set
        {
            this._collection.ConvertRichtext();

            if (value)
            {
                _ = this.CreateNode(STRIKE_PATH);
            }
            else
            {
                this.DeleteNode(STRIKE_PATH);
            }

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    const string UNDERLINE_PATH = "d:rPr/d:u";

    /// <summary>
    /// Underlined text
    /// </summary>
    public bool UnderLine
    {
        get { return this.ExistsNode(UNDERLINE_PATH); }
        set
        {
            this._collection.ConvertRichtext();

            if (value)
            {
                _ = this.CreateNode(UNDERLINE_PATH);
            }
            else
            {
                this.DeleteNode(UNDERLINE_PATH);
            }

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    const string VERT_ALIGN_PATH = "d:rPr/d:vertAlign/@val";

    /// <summary>
    /// Vertical Alignment
    /// </summary>
    public ExcelVerticalAlignmentFont VerticalAlign
    {
        get
        {
            string v = this.GetXmlNodeString(VERT_ALIGN_PATH);

            if (v == "")
            {
                return ExcelVerticalAlignmentFont.None;
            }
            else
            {
                try
                {
                    return (ExcelVerticalAlignmentFont)Enum.Parse(typeof(ExcelVerticalAlignmentFont), v, true);
                }
                catch
                {
                    return ExcelVerticalAlignmentFont.None;
                }
            }
        }
        set
        {
            this._collection.ConvertRichtext();

            if (value == ExcelVerticalAlignmentFont.None)
            {
                // If Excel 2010 encounters a vertical align value of blank, it will not load
                // the spreadsheet. So if None is specified, delete the node, it will be 
                // recreated if a new value is applied later.
                this.DeleteNode(VERT_ALIGN_PATH, true);
            }
            else
            {
                this.SetXmlNodeString(VERT_ALIGN_PATH, value.ToString().ToLowerInvariant());
            }

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    const string SIZE_PATH = "d:rPr/d:sz/@val";

    /// <summary>
    /// Font size
    /// </summary>
    public float Size
    {
        get { return Convert.ToSingle(this.GetXmlNodeDecimal(SIZE_PATH)); }
        set
        {
            this._collection.ConvertRichtext();
            this.SetXmlNodeString(SIZE_PATH, value.ToString(CultureInfo.InvariantCulture));

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    const string FONT_PATH = "d:rPr/d:rFont/@val";

    /// <summary>
    /// Name of the font
    /// </summary>
    public string FontName
    {
        get { return this.GetXmlNodeString(FONT_PATH); }
        set
        {
            this._collection.ConvertRichtext();
            this.SetXmlNodeString(FONT_PATH, value);

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    internal const string COLOR_PATH = "d:rPr/d:color/@rgb";
    internal const string COLOR_THEME_PATH = "d:rPr/d:color/@theme";
    internal const string COLOR_TINT_PATH = "d:rPr/d:color/@tint";
    internal const string RPR_PATH = "d:rPr";

    /// <summary>
    /// Text color.
    /// Also see <seealso cref="ColorSettings"/>
    /// </summary>
    public Color Color
    {
        get
        {
            string? col = this.GetXmlNodeString(COLOR_PATH);
            double? tint = this.GetXmlNodeDoubleNull(COLOR_TINT_PATH);
            Color ret;

            if (col == "")
            {
                int? v = this.GetXmlNodeIntNull(COLOR_THEME_PATH);

                if (v.HasValue)
                {
                    ret = Utils.ColorConverter.GetThemeColor(this._collection._ws.Workbook.ThemeManager.GetOrCreateTheme(), (eThemeSchemeColor)v);
                }
                else
                {
                    if (this._collection._cells == null || this.ExistsNode(RPR_PATH)) //If the rPn element exist but no color element, automatic should be used.
                    {
                        ret = Utils.ColorConverter.GetThemeColor(this._collection._ws.Workbook.ThemeManager.GetOrCreateTheme().ColorScheme.Dark1);
                    }
                    else
                    {
                        //If not color is set, return the font of the first cell in the range.
                        string? s = this._collection._cells.Style.Font.Color.LookupColor();
                        ret = Color.FromArgb(int.Parse(s.Substring(1), NumberStyles.AllowHexSpecifier));
                    }
                }
            }
            else
            {
                ret = Color.FromArgb(int.Parse(col, NumberStyles.AllowHexSpecifier));
            }

            if (tint.HasValue)
            {
                return Utils.ColorConverter.ApplyTint(ret, tint.Value);
            }

            return ret;
        }
        set
        {
            this._collection.ConvertRichtext();
            this.SetXmlNodeString(COLOR_PATH, value.ToArgb().ToString("X") /*.Substring(2, 6)*/);

            if (this._callback != null)
            {
                this._callback();
            }
        }
    }

    ExcelRichTextColor _colorSettings;

    /// <summary>
    /// Color settings.
    /// <seealso cref="Color"/>
    /// </summary>
    public ExcelRichTextColor ColorSettings
    {
        get { return this._colorSettings ??= new ExcelRichTextColor(this.NameSpaceManager, this.TopNode, this); }
    }

    /// <summary>
    /// A referens to the richtext collection
    /// </summary>
    public ExcelRichTextCollection _collection { get; set; }
}