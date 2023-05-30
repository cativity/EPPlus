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
using System.Xml;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Style;

/// <summary>
/// Collection of Richtext objects
/// </summary>
public class ExcelRichTextCollection : XmlHelper, IEnumerable<ExcelRichText>
{
    List<ExcelRichText> _list = new List<ExcelRichText>();
    internal ExcelRangeBase _cells;
    internal ExcelWorksheet _ws;

    internal ExcelRichTextCollection(XmlNamespaceManager ns, XmlNode topNode, ExcelWorksheet ws)
        : base(ns, topNode)
    {
        XmlNodeList? nl = topNode.SelectNodes("d:r", this.NameSpaceManager);

        if (nl != null)
        {
            foreach (XmlNode n in nl)
            {
                this._list.Add(new ExcelRichText(ns, n, this));
            }
        }

        this._ws = ws;
    }

    internal ExcelRichTextCollection(XmlNamespaceManager ns, XmlNode topNode, ExcelRangeBase cells)
        : this(ns, topNode, cells._worksheet) =>
        this._cells = cells;

    /// <summary>
    /// Collection containing the richtext objects
    /// </summary>
    /// <param name="Index"></param>
    /// <returns></returns>
    public ExcelRichText this[int Index]
    {
        get
        {
            ExcelRichText? item = this._list[Index];

            if (this._cells != null)
            {
                item.SetCallback(this.UpdateCells);
            }

            return item;
        }
    }

    /// <summary>
    /// Items in the list
    /// </summary>
    public int Count => this._list.Count;

    /// <summary>
    /// Add a rich text string
    /// </summary>
    /// <param name="Text">The text to add</param>
    /// <param name="NewParagraph">Adds a new paragraph before text. This will add a new line break.</param>
    /// <returns></returns>
    public ExcelRichText Add(string Text, bool NewParagraph = false)
    {
        if (NewParagraph)
        {
            Text += "\n";
        }

        return this.Insert(this._list.Count, Text);
    }

    /// <summary>
    /// Insert a rich text string at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index at which rich text should be inserted.</param>
    /// <param name="text">The text to insert.</param>
    /// <returns></returns>
    public ExcelRichText Insert(int index, string text)
    {
        if (text == null)
        {
            throw new ArgumentException("Text can't be null", nameof(text));
        }

        this.ConvertRichtext();
        XmlDocument doc;

        if (this.TopNode is XmlDocument)
        {
            doc = this.TopNode as XmlDocument;
        }
        else
        {
            doc = this.TopNode.OwnerDocument;
        }

        XmlElement? node = doc.CreateElement("d", "r", ExcelPackage.schemaMain);

        if (index < this._list.Count)
        {
            _ = this.TopNode.InsertBefore(node, this.TopNode.ChildNodes[index]);
        }
        else
        {
            _ = this.TopNode.AppendChild(node);
        }

        ExcelRichText? rt = new ExcelRichText(this.NameSpaceManager, node, this);

        if (this._list.Count > 0)
        {
            ExcelRichText prevItem = this._list[index < this._list.Count ? index : this._list.Count - 1];
            rt.FontName = prevItem.FontName;
            rt.Size = prevItem.Size;

            if (prevItem.Color.IsEmpty)
            {
                rt.Color = Color.Black;
            }
            else
            {
                rt.Color = prevItem.Color;
            }

            rt.PreserveSpace = rt.PreserveSpace;
            rt.Bold = prevItem.Bold;
            rt.Italic = prevItem.Italic;
            rt.UnderLine = prevItem.UnderLine;
        }
        else if (this._cells == null)
        {
            rt.FontName = "Calibri";
            rt.Size = 11;
        }
        else
        {
            ExcelStyle? style = this._cells.Offset(0, 0).Style;
            rt.FontName = style.Font.Name;
            rt.Size = style.Font.Size;
            rt.Bold = style.Font.Bold;
            rt.Italic = style.Font.Italic;
            this._cells.SetIsRichTextFlag(true);
        }

        rt.Text = text;
        rt.PreserveSpace = true;

        if (this._cells != null)
        {
            rt.SetCallback(this.UpdateCells);
            this.UpdateCells();
        }

        this._list.Insert(index, rt);

        return rt;
    }

    internal void ConvertRichtext()
    {
        if (this._cells == null)
        {
            return;
        }

        bool isRt = this._cells.Worksheet._flags.GetFlagValue(this._cells._fromRow, this._cells._fromCol, CellFlags.RichText);

        if (this.Count == 1 && isRt == false)
        {
            this._cells.Worksheet._flags.SetFlagValue(this._cells._fromRow, this._cells._fromCol, true, CellFlags.RichText);
            int s = this._cells.Worksheet.GetStyleInner(this._cells._fromRow, this._cells._fromCol);

            //var fnt = cell.Style.Font;
            ExcelFont? fnt = this._cells.Worksheet.Workbook.Styles
                                 .GetStyleObject(s, this._cells.Worksheet.PositionId, ExcelCellBase.GetAddress(this._cells._fromRow, this._cells._fromCol))
                                 .Font;

            this[0].PreserveSpace = true;
            this[0].Bold = fnt.Bold;
            this[0].FontName = fnt.Name;
            this[0].Italic = fnt.Italic;
            this[0].Size = fnt.Size;
            this[0].UnderLine = fnt.UnderLine;

            if (fnt.Color.Rgb != "" && int.TryParse(fnt.Color.Rgb, NumberStyles.HexNumber, null, out int hex))
            {
                this[0].Color = Color.FromArgb(hex);
            }
        }
    }

    internal void UpdateCells() => this._cells.SetValueRichText(this.TopNode.InnerXml);

    /// <summary>
    /// Clear the collection
    /// </summary>
    public void Clear()
    {
        this._list.Clear();
        this.TopNode.RemoveAll();

        if (this._cells != null)
        {
            this._cells.DeleteMe(this._cells, false, true, true, true, false, true, false, false, false);
            this._cells.SetIsRichTextFlag(false);
        }
    }

    /// <summary>
    /// Removes an item at the specific index
    /// </summary>
    /// <param name="Index"></param>
    public void RemoveAt(int Index)
    {
        _ = this.TopNode.RemoveChild(this._list[Index].TopNode);
        this._list.RemoveAt(Index);

        if (this._cells != null && this._list.Count == 0)
        {
            this._cells.SetIsRichTextFlag(false);
        }
    }

    /// <summary>
    /// Removes an item
    /// </summary>
    /// <param name="Item"></param>
    public void Remove(ExcelRichText Item)
    {
        _ = this.TopNode.RemoveChild(Item.TopNode);
        _ = this._list.Remove(Item);
        this.UpdateCells();

        if (this._cells != null && this._list.Count == 0)
        {
            this._cells.SetIsRichTextFlag(false);
        }
    }

    /// <summary>
    /// The text
    /// </summary>
    public string Text
    {
        get
        {
            StringBuilder sb = new StringBuilder();

            foreach (ExcelRichText? item in this._list)
            {
                _ = sb.Append(item.Text);
            }

            return sb.ToString();
        }
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                this.Clear();
            }
            else if (this.Count == 0)
            {
                _ = this.Add(value);
            }
            else
            {
                this[0].Text = value;

                for (int ix = 1; ix < this.Count; ix++)
                {
                    this.RemoveAt(ix);
                }
            }
        }
    }

    /// <summary>
    /// Returns the rich text as a html string.
    /// </summary>
    public string HtmlText
    {
        get
        {
            StringBuilder? sb = new StringBuilder();

            foreach (ExcelRichText? item in this._list)
            {
                item.WriteHtmlText(sb);
            }

            return sb.ToString();
        }
    }

    #region IEnumerable<ExcelRichText> Members

    IEnumerator<ExcelRichText> IEnumerable<ExcelRichText>.GetEnumerator() =>
        this._list.Select(x =>
            {
                x.SetCallback(this.UpdateCells);

                return x;
            })
            .GetEnumerator();

    #endregion

    #region IEnumerable Members

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this._list.Select(x =>
            {
                x.SetCallback(this.UpdateCells);

                return x;
            })
            .GetEnumerator();

    #endregion
}