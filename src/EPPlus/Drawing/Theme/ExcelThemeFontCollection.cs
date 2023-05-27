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
using OfficeOpenXml.Drawing.Style.Font;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme;

/// <summary>
/// A collection of fonts in a theme
/// </summary>
public class ExcelThemeFontCollection : XmlHelper, IEnumerable<ExcelDrawingFontBase>
{
    List<ExcelDrawingFontBase> _lst = new List<ExcelDrawingFontBase>();
    ExcelPackage _pck;

    internal ExcelThemeFontCollection(ExcelPackage pck, XmlNamespaceManager nameSpaceManager, XmlNode topNode)
        : base(nameSpaceManager, topNode)
    {
        this._pck = pck;

        foreach (XmlNode node in topNode.ChildNodes)
        {
            if (node.LocalName == "font")
            {
                this._lst.Add(new ExcelDrawingFont(nameSpaceManager, node));
            }
            else
            {
                this._lst.Add(new ExcelDrawingFontSpecial(nameSpaceManager, node));
            }
        }
    }

    /// <summary>
    /// The collection index
    /// </summary>
    /// <param name="index">The index</param>
    /// <returns></returns>
    public ExcelDrawingFontBase this[int index]
    {
        get { return this._lst[index]; }
    }

    /// <summary>
    /// Adds a normal font to the collection
    /// </summary>
    /// <param name="typeface">The typeface, or name of the font</param>
    /// <param name="script">The script, or language, in which the typeface is supposed to be used</param>
    /// <returns>The font</returns>
    public ExcelDrawingFont Add(string typeface, string script)
    {
        XmlNode e = this.TopNode.OwnerDocument.CreateElement("a", "font", ExcelPackage.schemaDrawings);
        _ = this.TopNode.AppendChild(e);
        ExcelDrawingFont? f = new ExcelDrawingFont(this.NameSpaceManager, e) { Typeface = typeface, Script = script };
        this._lst.Add(f);

        return f;
    }

    /// <summary>
    /// Removes the item from the collection
    /// </summary>
    /// <param name="index">The index of the item to remove</param>
    public void RemoveAt(int index)
    {
        if (index < 0 || index >= this._lst.Count)
        {
            throw new IndexOutOfRangeException();
        }

        this.Remove(this._lst[index]);
    }

    /// <summary>
    /// Removes the item from the collection
    /// </summary>
    /// <param name="item">The item to remove</param>
    public void Remove(ExcelDrawingFontBase item)
    {
        if (item is ExcelDrawingFontSpecial)
        {
            throw new InvalidOperationException("Cant remove this type of font.");
        }

        _ = item.TopNode.ParentNode.RemoveChild(item.TopNode);
        _ = this._lst.Remove(item);
    }

    /// <summary>
    /// Set the latin font of the collection
    /// </summary>
    /// <param name="typeface">The typeface, or name of the font</param>
    public void SetLatinFont(string typeface)
    {
        if (this._pck.Workbook.Styles.Fonts.Count > 0 && string.IsNullOrEmpty(typeface) == false)
        {
            string? id = this._pck.Workbook.Styles.Fonts[0].Id;
            this._pck.Workbook.Styles.Fonts[0].Name = typeface;
            _ = this._pck.Workbook.Styles.Fonts._dic.Remove(id);
            this._pck.Workbook.Styles.Fonts._dic.Add(this._pck.Workbook.Styles.Fonts[0].Id, 0);
        }

        this.SetSpecialFont(typeface, eFontType.Latin);
    }

    /// <summary>
    /// Set the complex font of the collection
    /// </summary>
    /// <param name="typeface">The typeface, or name of the font</param>
    public void SetComplexFont(string typeface)
    {
        this.SetSpecialFont(typeface, eFontType.Complex);
    }

    /// <summary>
    /// Set the East Asian font of the collection
    /// </summary>
    /// <param name="typeface">The typeface, or name of the font</param>
    public void SetEastAsianFont(string typeface)
    {
        this.SetSpecialFont(typeface, eFontType.EastAsian);
    }

    private void SetSpecialFont(string typeface, eFontType fontType)
    {
        ExcelDrawingFontBase? f = this._lst.Where(x => x is ExcelDrawingFontSpecial sf && sf.Type == fontType).FirstOrDefault()
                                  ?? this.AddSpecialFont(fontType, typeface);

        f.Typeface = typeface;
    }

    /// <summary>
    /// Adds a special font to the fonts collection
    /// </summary>
    /// <param name="type">The font type</param>
    /// <param name="typeface">The typeface, or name of the font</param>
    /// <returns>The font</returns>
    public ExcelDrawingFontSpecial AddSpecialFont(eFontType type, string typeface)
    {
        string typeName;

        switch (type)
        {
            case eFontType.Complex:
                typeName = "cs";

                break;

            case eFontType.EastAsian:
                typeName = "ea";

                break;

            case eFontType.Latin:
                typeName = "latin";

                break;

            case eFontType.Symbol:
                typeName = "sym";

                break;

            default:
                throw new ArgumentException("Please use the Add method to add normal fonts");
        }

        XmlNode e = this.TopNode.OwnerDocument.CreateElement("a", typeName, ExcelPackage.schemaDrawings);
        _ = this.TopNode.AppendChild(e);
        ExcelDrawingFontSpecial? f = new ExcelDrawingFontSpecial(this.NameSpaceManager, e) { Typeface = typeface };
        this._lst.Add(f);

        return f;
    }

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get { return this._lst.Count; }
    }

    /// <summary>
    /// Gets an enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelDrawingFontBase> GetEnumerator()
    {
        return this._lst.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._lst.GetEnumerator();
    }
}