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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme;

/// <summary>
/// Defines the line styles within the theme
/// </summary>
public class ExcelThemeLineStyles : XmlHelper, IEnumerable<ExcelThemeLine>
{
    List<ExcelThemeLine> _list;
    internal ExcelThemeLineStyles(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
    {
        this._list = new List<ExcelThemeLine>();
        foreach (XmlNode node in topNode.ChildNodes)
        {
            this._list.Add(new ExcelThemeLine(nameSpaceManager, node));
        }
    }
    /// <summary>
    /// Gets the enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelThemeLine> GetEnumerator()
    {
        return this._list.GetEnumerator();
    }
    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._list.GetEnumerator();
    }
    /// <summary>
    /// Indexer for the collection
    /// </summary>
    /// <param name="index">The index</param>
    /// <returns>The line style</returns>
    public ExcelThemeLine this[int index]
    {
        get
        {
            return this._list[index];
        }
    }
    /// <summary>
    /// Adds a new line to the collection
    /// </summary>
    /// <returns>The line</returns>
    public ExcelThemeLine Add()
    {
        XmlElement? node = this.TopNode.OwnerDocument.CreateElement("a", "ln", ExcelPackage.schemaMain);
        this.TopNode.AppendChild(node);
        return new ExcelThemeLine(this.NameSpaceManager, this.TopNode);
    }
    /// <summary>
    /// Removes a line item from the collection
    /// </summary>
    /// <param name="item">The item</param>
    public void Remove(ExcelThemeLine item)
    {
        if (this._list.Count == 3)
        {
            throw new InvalidOperationException("Collection must contain at least 3 items");
        }

        if (this._list.Contains(item))
        {
            this._list.Remove(item);
            item.TopNode.ParentNode.RemoveChild(item.TopNode);
        }
    }
    /// <summary>
    /// Remove the line style at the specified index. The collection must have at least three line styles.
    /// </summary>
    /// <param name="Index">The index</param>
    public void Remove(int Index)
    {
        if (this._list.Count == 3)
        {
            throw new InvalidOperationException("Collection must contain at least 3 items");
        }

        if (Index >= this._list.Count)
        {
            throw new ArgumentException("Index", "Index out of range");
        }

        this._list.Remove(this._list[Index]);
    }
    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get
        {
            return this._list.Count;
        }
    }
}