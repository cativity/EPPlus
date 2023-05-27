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
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill;

/// <summary>
/// A collection of colors and their positions used for a gradiant fill.
/// </summary>
public class ExcelDrawingGradientFillColorList : IEnumerable<ExcelDrawingGradientFillColor>
{
    List<ExcelDrawingGradientFillColor> _lst = new List<ExcelDrawingGradientFillColor>();
    private XmlNamespaceManager _nsm;
    private XmlNode _topNode;
    private XmlNode _gsLst=null;
    private string _path;
    private string[] _schemaNodeOrder;
    internal ExcelDrawingGradientFillColorList(XmlNamespaceManager nsm, XmlNode topNode, string path, string[] schemaNodeOrder)
    {
        this._nsm = nsm;
        this._topNode = topNode;
        this._path = path;
        this._schemaNodeOrder = schemaNodeOrder;
    }
    /// <summary>
    /// Indexer for the collection
    /// </summary>
    /// <param name="index">The index in the collection</param>
    /// <returns>The color</returns>
    public ExcelDrawingGradientFillColor this[int index]
    {
        get
        {
            return this._lst[index];
        }
    }
    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get
        {
            return this._lst.Count;
        }
    }
    /// <summary>
    /// Gets the first occurance with the color with the specified position
    /// </summary>
    /// <param name="position">The position in percentage</param>
    /// <returns>The color</returns>
    public ExcelDrawingGradientFillColor this[double position]
    {
        get
        {
            return this._lst.Find(i => i.Position == position);
        }
    }
    /// <summary>
    /// Adds a RGB color at the specified position
    /// </summary>
    /// <param name="position">The position</param>
    /// <param name="color">The Color</param>
    public void AddRgb(double position, Color color)
    {
        ExcelDrawingGradientFillColor? gs = this.GetGradientFillColor(position);
        gs.Color.SetRgbColor(color);
        this._lst.Add(gs);
    }
    /// <summary>
    /// Adds a RGB percentage color at the specified position
    /// </summary>
    /// <param name="position">The position</param>
    /// <param name="redPercentage">The percentage of red</param>
    /// <param name="greenPercentage">The percentage of green</param>
    /// <param name="bluePercentage">The percentage of blue</param>
    public void AddRgbPercentage(double position, double redPercentage, double greenPercentage, double bluePercentage)
    {
        ExcelDrawingGradientFillColor? gs = this.GetGradientFillColor(position);
        gs.Color.SetRgbPercentageColor(redPercentage, greenPercentage, bluePercentage);
        this._lst.Add(gs);
    }
    /// <summary>
    /// Adds a theme color at the specified position
    /// </summary>
    /// <param name="position">The position</param>
    /// <param name="color">The theme color</param>
    public void AddScheme(double position, eSchemeColor color)
    {
        ExcelDrawingGradientFillColor? gs = this.GetGradientFillColor(position);
        gs.Color.SetSchemeColor(color);
        this._lst.Add(gs);
    }
    /// <summary>
    /// Adds a system color at the specified position
    /// </summary>
    /// <param name="position">The position</param>
    /// <param name="color">The system color</param>
    public void AddSystem(double position, eSystemColor color)
    {
        ExcelDrawingGradientFillColor? gs = this.GetGradientFillColor(position);
        gs.Color.SetSystemColor(color);
        this._lst.Add(gs);
    }
    /// <summary>
    /// Adds a HSL color at the specified position
    /// </summary>
    /// <param name="position">The position</param>
    /// <param name="hue">The hue part. Ranges from 0-360</param>
    /// <param name="saturation">The saturation part. Percentage</param>
    /// <param name="luminance">The luminance part. Percentage</param>
    public void AddHsl(double position, double hue, double saturation, double luminance)
    {            
        ExcelDrawingGradientFillColor? gs = this.GetGradientFillColor(position);
        gs.Color.SetHslColor(hue, saturation, luminance);
        this._lst.Add(gs);
    }
    /// <summary>
    /// Adds a HSL color at the specified position
    /// </summary>
    /// <param name="position">The position</param>
    /// <param name="color">The preset color</param>
    public void AddPreset(double position, ePresetColor color)
    {
        ExcelDrawingGradientFillColor? gs = this.GetGradientFillColor(position);
        gs.Color.SetPresetColor(color);
        this._lst.Add(gs);
    }

    private ExcelDrawingGradientFillColor GetGradientFillColor(double position)
    {
        if (position < 0 || position > 100)
        {
            throw new ArgumentOutOfRangeException("Position must be between 0 and 100");
        }
        XmlNode node = null;
        for (int i = 0; i < this._lst.Count; i++)
        {
            if (this._lst[i].Position > position)
            {
                node = this.AddGs(position, this._lst[i].TopNode);
            }
        }
        node = this.AddGs(position, null);

        ExcelDrawingGradientFillColor? tc = new ExcelDrawingGradientFillColor()
        {
            Position = position,
            Color = new ExcelDrawingColorManager(this._nsm, node, "", this._schemaNodeOrder),
            TopNode = node
        };
        return tc;
    }

    private XmlElement AddGs(double position, XmlNode node)
    {
        if(this._gsLst==null)
        {
            XmlHelper? xml = XmlHelperFactory.Create(this._nsm, this._topNode);
            this._gsLst=xml.CreateNode(this._path);
        }
        XmlElement? gs = this._gsLst.OwnerDocument.CreateElement("a", "gs", ExcelPackage.schemaDrawings);
        if (node == null)
        {
            this._gsLst.AppendChild(gs);
        }
        else
        {
            this._gsLst.InsertBefore(gs, node);
        }
        gs.SetAttribute("pos", (position * 1000).ToString());
        return gs;
    }
    /// <summary>
    /// Gets the enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelDrawingGradientFillColor> GetEnumerator()
    {
        return this._lst.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._lst.GetEnumerator();
    }

    internal void Add(double position, XmlNode node)
    {
        this._lst.Add(new ExcelDrawingGradientFillColor()
        {
            Position = position,
            Color = new ExcelDrawingColorManager(this._nsm, node, "", this._schemaNodeOrder),
            TopNode = node
        });
    }
}