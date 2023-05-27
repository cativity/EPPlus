/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/29/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// A collection of colors and their positions used for a gradiant fill.
/// </summary>
public class ExcelDxfGradientFillColorCollection : DxfStyleBase, IEnumerable<ExcelDxfGradientFillColor>
{
    List<ExcelDxfGradientFillColor> _lst = new List<ExcelDxfGradientFillColor>();

    internal ExcelDxfGradientFillColorCollection(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback)
    {
    }

    /// <summary>
    /// Get the enumerator
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelDxfGradientFillColor> GetEnumerator()
    {
        return this._lst.GetEnumerator();
    }

    /// <summary>
    /// Get the enumerator
    /// </summary>
    /// <returns>The enumerator</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._lst.GetEnumerator();
    }

    /// <summary>
    /// Indexer for the collection
    /// </summary>
    /// <param name="index">The index in the collection</param>
    /// <returns>The color</returns>
    public ExcelDxfGradientFillColor this[int index]
    {
        get { return this._lst[index]; }
    }

    /// <summary>
    /// Gets the first occurance with the color with the specified position
    /// </summary>
    /// <param name="position">The position in percentage</param>
    /// <returns>The color</returns>
    public ExcelDxfGradientFillColor this[double position]
    {
        get { return this._lst.Find(i => i.Position == position); }
    }

    /// <summary>
    /// Adds a RGB color at the specified position
    /// </summary>
    /// <param name="position">The position from 0 to 100%</param>
    /// <returns>The gradient color position object</returns>
    public ExcelDxfGradientFillColor Add(double position)
    {
        if (position < 0 && position > 100)
        {
            throw new ArgumentOutOfRangeException("position", "Must be a value between 0 and 100");
        }

        ExcelDxfGradientFillColor? color = new ExcelDxfGradientFillColor(this._styles, position, this._callback);
        color.Color.Auto = true;
        this._lst.Add(color);

        return color;
    }

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get { return this._lst.Count; }
    }

    internal override string Id
    {
        get
        {
            string? id = "";

            foreach (ExcelDxfGradientFillColor? c in this._lst.OrderBy(x => x.Position))
            {
                id += c.Id;
            }

            return id;
        }
    }

    /// <summary>
    /// If the style has any value set
    /// </summary>
    public override bool HasValue
    {
        get { return this._lst.Count > 0; }
    }

    /// <summary>
    /// Remove the style at the index in the collection.
    /// </summary>
    /// <param name="index"></param>
    public void RemoveAt(int index)
    {
        this._lst.RemoveAt(index);
    }

    /// <summary>
    /// Remove the style at the position from the collection.
    /// </summary>
    /// <param name="position"></param>
    public void RemoveAt(double position)
    {
        ExcelDxfGradientFillColor? item = this._lst.Find(i => i.Position == position);

        if (item != null)
        {
            _ = this._lst.Remove(item);
        }
    }

    /// <summary>
    /// Remove the style from the collection
    /// </summary>
    /// <param name="item"></param>
    public void Remove(ExcelDxfGradientFillColor item)
    {
        _ = this._lst.Remove(item);
    }

    /// <summary>
    /// Clear all style items from the collection
    /// </summary>
    public override void Clear()
    {
        this._lst.Clear();
    }

    internal override void CreateNodes(XmlHelper helper, string path)
    {
        if (this._lst.Count > 0)
        {
            foreach (ExcelDxfGradientFillColor? c in this._lst)
            {
                c.CreateNodes(helper, path);
            }
        }
    }

    internal override void SetStyle()
    {
        if (this._callback != null && this._lst.Count > 0)
        {
            foreach (ExcelDxfGradientFillColor? c in this._lst)
            {
                c.SetStyle();
            }
        }
    }

    internal override DxfStyleBase Clone()
    {
        ExcelDxfGradientFillColorCollection? ret = new ExcelDxfGradientFillColorCollection(this._styles, this._callback);

        foreach (ExcelDxfGradientFillColor? c in this._lst)
        {
            ret._lst.Add((ExcelDxfGradientFillColor)c.Clone());
        }

        return ret;
    }
}