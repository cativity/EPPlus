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
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.Linq;
using System.Globalization;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style;

/// <summary>
/// A collection of Paragraph objects
/// </summary>
public class ExcelParagraphCollection : XmlHelper, IEnumerable<ExcelParagraph>
{
    List<ExcelParagraph> _list = new List<ExcelParagraph>();
    private readonly ExcelDrawing _drawing;
    private readonly string _path;
    private readonly List<XmlElement> _paragraphs = new List<XmlElement>();
    private readonly float _defaultFontSize;

    internal ExcelParagraphCollection(ExcelDrawing drawing,
                                      XmlNamespaceManager ns,
                                      XmlNode topNode,
                                      string path,
                                      string[] schemaNodeOrder,
                                      float defaultFontSize = 11)
        : base(ns, topNode)
    {
        this._drawing = drawing;
        this._defaultFontSize = defaultFontSize;

        this.AddSchemaNodeOrder(schemaNodeOrder,
                                new string[]
                                {
                                    "strRef", "rich", "f", "strCache", "bodyPr", "lstStyle", "p", "ptCount", "pt", "pPr", "lnSpc", "spcBef", "spcAft",
                                    "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip",
                                    "tabLst", "defRPr", "r", "br", "fld", "endParaRPr"
                                });

        this._path = path;
        XmlNodeList? pars = this.TopNode.SelectNodes(path, this.NameSpaceManager);

        foreach (XmlElement par in pars)
        {
            this._paragraphs.Add(par);
            XmlNodeList? nl = par.SelectNodes("a:r", this.NameSpaceManager);

            if (nl != null)
            {
                foreach (XmlNode n in nl)
                {
                    if (this._list.Count == 0 || n.ParentNode != this._list[this._list.Count - 1].TopNode.ParentNode)
                    {
                        this._paragraphs.Add((XmlElement)n.ParentNode);
                    }

                    this._list.Add(new ExcelParagraph(drawing._drawings, ns, n, "", schemaNodeOrder));
                }
            }
        }
    }

    /// <summary>
    /// The indexer for this collection
    /// </summary>
    /// <param name="Index">The index</param>
    /// <returns></returns>
    public ExcelParagraph this[int Index]
    {
        get { return this._list[Index]; }
    }

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get { return this._list.Count; }
    }

    /// <summary>
    /// Add a rich text string
    /// </summary>
    /// <param name="Text">The text to add</param>
    /// <param name="NewParagraph">This will be a new line. Is ignored for first item added to the collection</param>
    /// <returns></returns>
    public ExcelParagraph Add(string Text, bool NewParagraph = false)
    {
        XmlDocument doc;

        if (this.TopNode is XmlDocument)
        {
            doc = this.TopNode as XmlDocument;
        }
        else
        {
            doc = this.TopNode.OwnerDocument;
        }

        XmlNode parentNode;

        if (NewParagraph && this._list.Count != 0)
        {
            parentNode = this.CreateNode(this._path, false, true);
            this._paragraphs.Add((XmlElement)parentNode);
            XmlNode? p = this._list[0].TopNode.ParentNode.ParentNode.SelectSingleNode("a:pPr", this.NameSpaceManager);

            if (p != null)
            {
                parentNode.InnerXml = p.OuterXml;
            }
        }
        else if (this._paragraphs.Count > 1)
        {
            parentNode = this._paragraphs[this._paragraphs.Count - 1];
        }
        else
        {
            parentNode = this.CreateNode(this._path);
            this._paragraphs.Add((XmlElement)parentNode);
            XmlNode? defNode = this.CreateNode(this._path + "/a:pPr/a:defRPr");

            if (defNode.InnerXml == "")
            {
                ((XmlElement)defNode).SetAttribute("sz", (this._defaultFontSize * 100).ToString(CultureInfo.InvariantCulture));
                ExcelNamedStyleXml? normalStyle = this._drawing._drawings.Worksheet.Workbook.Styles.GetNormalStyle();

                if (normalStyle == null)
                {
                    defNode.InnerXml = "<a:latin typeface=\"Calibri\" /><a:cs typeface=\"Calibri\" />";
                }
                else
                {
                    defNode.InnerXml = $"<a:latin typeface=\"{normalStyle.Style.Font.Name}\"/><a:cs typeface=\"{normalStyle.Style.Font.Name}\"/>";
                }
            }
        }

        XmlElement? node = doc.CreateElement("a", "r", ExcelPackage.schemaDrawings);
        parentNode.AppendChild(node);
        XmlElement? childNode = doc.CreateElement("a", "rPr", ExcelPackage.schemaDrawings);
        node.AppendChild(childNode);
        ExcelParagraph? rt = new ExcelParagraph(this._drawing._drawings, this.NameSpaceManager, node, "", this.SchemaNodeOrder);

        //var normalStyle = _drawing._drawings.Worksheet.Workbook.Styles.GetNormalStyle();
        //if (normalStyle == null)
        //{
        //    rt.ComplexFont = "Calibri";
        //    rt.LatinFont = "Calibri";
        //}
        //else
        //{
        //    rt.LatinFont = normalStyle.Style.Font.Name;
        //    rt.ComplexFont = normalStyle.Style.Font.Name;
        //}
        //rt.Size = _defaultFontSize;

        rt.Text = Text;
        this._list.Add(rt);

        return rt;
    }

    /// <summary>
    /// Removes all items in the collection
    /// </summary>
    public void Clear()
    {
        for (int ix = 0; ix < this._paragraphs.Count; ix++)
        {
            this._paragraphs[ix].ParentNode?.RemoveChild(this._paragraphs[ix]);
        }

        this._list.Clear();
        this._paragraphs.Clear();
    }

    /// <summary>
    /// Remove the item at the specified index
    /// </summary>
    /// <param name="Index">The index</param>
    public void RemoveAt(int Index)
    {
        XmlNode? node = this._list[Index].TopNode;

        while (node != null && node.Name != "a:r")
        {
            node = node.ParentNode;
        }

        node.ParentNode.RemoveChild(node);
        this._list.RemoveAt(Index);
    }

    /// <summary>
    /// Remove the specified item
    /// </summary>
    /// <param name="Item">The item</param>
    public void Remove(ExcelRichText Item)
    {
        this.TopNode.RemoveChild(Item.TopNode);
    }

    /// <summary>
    /// The full text 
    /// </summary>
    public string Text
    {
        get
        {
            StringBuilder sb = new StringBuilder();

            foreach (ExcelParagraph? item in this._list)
            {
                if (item.IsLastInParagraph)
                {
                    sb.AppendLine(item.Text);
                }
                else
                {
                    sb.Append(item.Text);
                }
            }

            if (sb.Length > 2)
            {
                sb.Remove(sb.Length - 2, 2); //Remove last crlf
            }

            return sb.ToString();
        }
        set
        {
            if (this.Count == 0)
            {
                this.Add(value);
            }
            else
            {
                this[0].Text = value;

                for (int ix = this._list.Count - 1; ix > 0; ix--)
                {
                    this.RemoveAt(ix);
                }
            }
        }
    }

    #region IEnumerable<ExcelRichText> Members

    IEnumerator<ExcelParagraph> IEnumerable<ExcelParagraph>.GetEnumerator()
    {
        return this._list.GetEnumerator();
    }

    #endregion

    #region IEnumerable Members

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return this._list.GetEnumerator();
    }

    #endregion
}