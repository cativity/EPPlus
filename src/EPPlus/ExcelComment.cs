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
using System.Linq;
using System.Text;
using OfficeOpenXml.Style;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;

namespace OfficeOpenXml;

/// <summary>
/// An Excel Cell Comment
/// </summary>
public class ExcelComment : ExcelVmlDrawingComment
{
    internal XmlHelper _commentHelper;
    private string _text;

    internal ExcelComment(XmlNamespaceManager ns, XmlNode commentTopNode, ExcelRangeBase cell)
        : base(null, cell, cell.Worksheet.VmlDrawings.NameSpaceManager)
    {
        //_commentHelper = new XmlHelper(ns, commentTopNode);
        this._commentHelper = XmlHelperFactory.Create(ns, commentTopNode);
        XmlNode? textElem = commentTopNode.SelectSingleNode("d:text", ns);

        if (textElem == null)
        {
            textElem = commentTopNode.OwnerDocument.CreateElement("text", ExcelPackage.schemaMain);
            _ = commentTopNode.AppendChild(textElem);
        }

        if (!cell.Worksheet._vmlDrawings.ContainsKey(cell.Start.Row, cell.Start.Column))
        {
            _ = cell.Worksheet._vmlDrawings.AddComment(cell);
        }

        this.TopNode = cell.Worksheet.VmlDrawings[cell.Start.Row, cell.Start.Column].TopNode;
        this.RichText = new ExcelRichTextCollection(ns, textElem, cell.Worksheet);
        XmlNode? tNode = textElem.SelectSingleNode("d:t", ns);

        if (tNode != null)
        {
            this._text = tNode.InnerText;
        }
    }

    const string AUTHORS_PATH = "d:comments/d:authors";
    const string AUTHOR_PATH = "d:comments/d:authors/d:author";

    /// <summary>
    /// The author
    /// </summary>
    public string Author
    {
        get
        {
            int authorRef = this._commentHelper.GetXmlNodeInt("@authorId");

            return this._commentHelper.TopNode.OwnerDocument
                       .SelectSingleNode(string.Format("{0}[{1}]", AUTHOR_PATH, authorRef + 1), this._commentHelper.NameSpaceManager)
                       .InnerText;
        }
        set
        {
            int authorRef = this.GetAuthor(value);
            this._commentHelper.SetXmlNodeString("@authorId", authorRef.ToString());
        }
    }

    private int GetAuthor(string value)
    {
        int authorRef = 0;
        bool found = false;

        foreach (XmlElement node in this._commentHelper.TopNode.OwnerDocument.SelectNodes(AUTHOR_PATH, this._commentHelper.NameSpaceManager))
        {
            if (node.InnerText == value)
            {
                found = true;

                break;
            }

            authorRef++;
        }

        if (!found)
        {
            XmlElement? elem = this._commentHelper.TopNode.OwnerDocument.CreateElement("d", "author", ExcelPackage.schemaMain);
            _ = this._commentHelper.TopNode.OwnerDocument.SelectSingleNode(AUTHORS_PATH, this._commentHelper.NameSpaceManager).AppendChild(elem);
            elem.InnerText = value;
        }

        return authorRef;
    }

    /// <summary>
    /// The comment text 
    /// </summary>
    public string Text
    {
        get
        {
            if (!string.IsNullOrEmpty(this.RichText.Text))
            {
                return this.RichText.Text;
            }

            return this._text;
        }
        set { this.RichText.Text = value; }
    }

    /// <summary>
    /// Sets the font of the first richtext item.
    /// </summary>
    public ExcelRichText Font
    {
        get
        {
            if (this.RichText.Count > 0)
            {
                return this.RichText[0];
            }

            return null;
        }
    }

    /// <summary>
    /// Richtext collection
    /// </summary>
    public ExcelRichTextCollection RichText { get; set; }

    /// <summary>
    /// Reference
    /// </summary>
    internal string Reference
    {
        get { return this._commentHelper.GetXmlNodeString("@ref"); }
        set
        {
            ExcelAddressBase? a = new ExcelAddressBase(value);
            int rows = a._fromRow - this.Range._fromRow;
            int cols = a._fromCol - this.Range._fromCol;
            this.Range.Address = value;
            this._commentHelper.SetXmlNodeString("@ref", value);

            this.From.Row += rows;
            this.To.Row += rows;

            this.From.Column += cols;
            this.To.Column += cols;

            this.Row = this.Range._fromRow - 1;
            this.Column = this.Range._fromCol - 1;
        }
    }
}