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

using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Coloring;

/// <summary>
/// Color transformation 
/// </summary>
public class ExcelColorTransformCollection : IEnumerable<IColorTransformItem>
{
    List<IColorTransformItem> _list = new List<IColorTransformItem>();
    XmlNamespaceManager _namespaceManager;
    XmlNode _topNode;

    /// <summary>
    /// For internal transformation calculations only. 
    /// </summary>
    internal ExcelColorTransformCollection()
    {
    }

    internal ExcelColorTransformCollection(XmlNamespaceManager nsm, XmlNode topNode)
    {
        this._namespaceManager = nsm;
        this._topNode = topNode;

        foreach (XmlElement e in topNode.ChildNodes)
        {
            eColorTransformType type = e.LocalName.ToEnum(eColorTransformType.Alpha);
            this._list.Add(new ExcelColorTransformItem(nsm, e, type));
        }
    }

    /// <summary>
    /// Indexer for the colletion
    /// </summary>
    /// <param name="index">The position in the list</param>
    /// <returns></returns>
    public IColorTransformItem this[int index] => this._list[index];

    /// <summary>
    /// Clear all items
    /// </summary>
    public void Clear()
    {
        foreach (IColorTransformItem? item in this._list)
        {
            if (item is ExcelColorTransformItem colorItem)
            {
                _ = colorItem.TopNode.ParentNode.RemoveChild(colorItem.TopNode);
            }
        }

        this._list.Clear();
    }

    /// <summary>
    /// Remote item at a specific position
    /// </summary>
    /// <param name="index">The postion in the list</param>
    public void RemoveAt(int index) => this.Remove(this._list[index]);

    /// <summary>
    /// Removes the specific item
    /// </summary>
    /// <param name="item">The item to remove</param>
    public void Remove(IColorTransformItem item)
    {
        if (item is ExcelColorTransformItem colorItem)
        {
            _ = colorItem.TopNode.ParentNode.RemoveChild(colorItem.TopNode);
        }

        _ = this._list.Remove(item);
    }

    /// <summary>
    /// Remove all items of a specific type
    /// </summary>
    /// <param name="type">The transformation type</param>
    public void RemoveOfType(eColorTransformType type)
    {
        for (int i = 0; i < this._list.Count; i++)
        {
            if (this._list[i].Type == type)
            {
                this._list.RemoveAt(i);
                i--;
            }
        }
    }

    #region Add methods

    #region Alpha

    /// <summary>
    /// The opacity as expressed by a percentage value
    /// Alpha equals 100-Transparancy
    /// </summary>
    /// <param name="value">The alpha value in percentage 0-100</param>
    public void AddAlpha(double value) => this.AddValue("alpha", eColorTransformType.Alpha, value);

    /// <summary>
    /// Specifies a more or less opaque version of its input color
    /// Alpha equals 100-Transparancy
    /// </summary>
    /// <param name="value">The alpha modulation in a positive percentage</param>
    public void AddAlphaModulation(double value) => this.AddValue("alphaMod", eColorTransformType.AlphaMod, value);

    /// <summary>
    /// Adds an alpha offset value. 
    /// </summary>
    /// <param name="value">The tint percentage. From 0-100</param>
    public void AddAlphaOffset(double value) => this.AddValue("alphaOff", eColorTransformType.AlphaOff, value);

    #endregion

    #region Hue

    /// <summary>
    /// Specifies the input color with the specified hue, but with its saturation and luminance unchanged
    /// </summary>
    /// <param name="value">The hue angle from 0-360</param>
    public void AddHue(double value) => this.AddValue("hue", eColorTransformType.Hue, value);

    /// <summary>
    /// Specifies the hue as expressed by a percentage relative to the input color
    /// </summary>
    /// <param name="value">The hue modulation in a positive percentage</param>
    public void AddHueModulation(double value) => this.AddValue("hueMod", eColorTransformType.HueMod, value);

    /// <summary>
    /// Specifies the actual angular value of the shift. The result of the shift shall be between 0 and 360 degrees.Shifts resulting in angular values less than 0 are treated as 0. 
    /// Shifts resulting in angular values greater than 360 are treated as 360.
    /// </summary>
    /// <param name="value">The hue offset value.</param>
    public void AddHueOffset(double value) => this.AddValue("hueOff", eColorTransformType.HueOff, value);

    #endregion

    #region Saturation

    /// <summary>
    /// Specifies the input color with the specified saturation, but with its hue and luminance unchanged
    /// </summary>
    /// <param name="value">The saturation percentage from 0-100</param>
    public void AddSaturation(double value) => this.AddValue("sat", eColorTransformType.Sat, value);

    /// <summary>
    /// Specifies the saturation as expressed by a percentage relative to the input color
    /// </summary>
    /// <param name="value">The saturation modulation in a positive percentage</param>
    public void AddSaturationModulation(double value) => this.AddValue("satMod", eColorTransformType.SatMod, value);

    /// <summary>
    /// Specifies the saturation as expressed by a percentage offset increase or decrease to the input color.
    /// Increases never increase the saturation beyond 100%, decreases never decrease the saturation below 0%.
    /// </summary>
    /// <param name="value">The saturation offset value</param>
    public void AddSaturationOffset(double value) => this.AddValue("satOff", eColorTransformType.SatOff, value);

    #endregion

    #region Luminance

    /// <summary>
    /// Specifies the input color with the specified luminance, but with its hue and saturation unchanged
    /// </summary>
    /// <param name="value">The luminance percentage from 0-100</param>
    public void AddLuminance(double value) => this.AddValue("lum", eColorTransformType.Lum, value);

    /// <summary>
    /// Specifies the luminance as expressed by a percentage relative to the input color
    /// </summary>
    /// <param name="value">The luminance modulation in a positive percentage</param>
    public void AddLuminanceModulation(double value) => this.AddValue("lumMod", eColorTransformType.LumMod, value);

    /// <summary>
    /// Specifies the luminance as expressed by a percentage offset increase or decrease to the input color.
    /// Increases never increase the luminance beyond 100%, decreases never decrease the saturation below 0%.
    /// </summary>
    /// <param name="value">The luminance offset value</param>
    public void AddLuminanceOffset(double value) => this.AddValue("lumOff", eColorTransformType.LumOff, value);

    #endregion

    #region Red

    /// <summary>
    /// Specifies the input color with the specific red component
    /// </summary>
    /// <param name="value">The red value</param>
    public void AddRed(double value) => this.AddValue("red", eColorTransformType.Red, value);

    /// <summary>
    /// Specifies the red component as expressed by a percentage relative to the input color component
    /// </summary>
    /// <param name="value">The red modulation value</param>
    public void AddRedModulation(double value) => this.AddValue("redMod", eColorTransformType.RedMod, value);

    /// <summary>
    /// Specifies the red component as expressed by a percentage offset increase or decrease to the input color component
    /// </summary>
    /// <param name="value">The red offset value.</param>
    public void AddRedOffset(double value) => this.AddValue("redOff", eColorTransformType.RedOff, value);

    #endregion

    #region Green

    /// <summary>
    /// Specifies the input color with the specific green component
    /// </summary>
    /// <param name="value">The green value</param>
    public void AddGreen(double value) => this.AddValue("green", eColorTransformType.Green, value);

    /// <summary>
    /// Specifies the green component as expressed by a percentage relative to the input color component
    /// </summary>
    /// <param name="value">The green modulation value</param>
    public void AddGreenModulation(double value) => this.AddValue("greenMod", eColorTransformType.GreenMod, value);

    /// <summary>
    /// Specifies the green component as expressed by a percentage offset increase or decrease to the input color component
    /// </summary>
    /// <param name="value">The green offset value.</param>
    public void AddGreenOffset(double value) => this.AddValue("greenOff", eColorTransformType.GreenOff, value);

    #endregion

    #region Blue

    /// <summary>
    /// Specifies the input color with the specific blue component
    /// </summary>
    /// <param name="value">The blue value</param>
    public void AddBlue(double value) => this.AddValue("blue", eColorTransformType.Blue, value);

    internal double FindValue(eColorTransformType alpha) => this._list.Find(x => x.Type == alpha)?.Value ?? 0;

    internal IColorTransformItem Find(eColorTransformType alpha) => this._list.Find(x => x.Type == alpha);

    /// <summary>
    /// Specifies the blue component as expressed by a percentage relative to the input color component
    /// </summary>
    /// <param name="value">The blue modulation value</param>
    public void AddBlueModulation(double value) => this.AddValue("blueMod", eColorTransformType.BlueMod, value);

    /// <summary>
    /// Specifies the blue component as expressed by a percentage offset increase or decrease to the input color component
    /// </summary>
    /// <param name="value">The blue offset value.</param>
    public void AddBlueOffset(double value) => this.AddValue("blueOff", eColorTransformType.BlueOff, value);

    #endregion

    /// <summary>
    /// Specifies a lighter version of its input color
    /// </summary>
    /// <param name="value">The tint value in percentage 0-100</param>
    public void AddTint(double value) => this.AddValue("tint", eColorTransformType.Tint, value);

    /// <summary>
    /// Specifies a lighter version of its input color
    /// </summary>
    /// <param name="value">The tint value in percentage 0-100</param>
    public void AddShade(double value) => this.AddValue("shade", eColorTransformType.Shade, value);

    #region Boolean Types

    /// <summary>
    /// Specifies that the color rendered should be the complement of its input color with the complement being defined as such.
    /// Two colors are called complementary if, when mixed they produce a shade of grey.For instance, the complement of red which is RGB (255, 0, 0) is cyan which is RGB(0, 255, 255)
    /// </summary>
    public void AddComplement() => this.AddValue("comp", eColorTransformType.Comp);

    /// <summary>
    /// Specifies that the output color rendered by the generating application should be the sRGB gamma shift of the input color.
    /// </summary>
    public void AddGamma() => this.AddValue("gamma", eColorTransformType.Gamma);

    /// <summary>
    /// Specifies a grayscale of its input color, taking into relative intensities of the red, green, and blue primaries.
    /// </summary>
    public void AddGray() => this.AddValue("gray", eColorTransformType.Gray);

    /// <summary>
    /// Specifies the inverse of its input color
    /// </summary>
    public void AddInverse() => this.AddValue("inv", eColorTransformType.Inv);

    /// <summary>
    /// Specifies that the output color rendered by the generating application should be the inverse sRGB gamma shift of the input color
    /// </summary>
    public void AddInverseGamma() => this.AddValue("invGamma", eColorTransformType.InvGamma);

    #endregion

    private void AddValue(string name, eColorTransformType type)
    {
        if (this._namespaceManager == null)
        {
            this._list.Add(new ExcelColorTransformSimpleItem() { Type = type });
        }
        else
        {
            XmlElement node = this.AddNode(name);
            this._list.Add(new ExcelColorTransformItem(this._namespaceManager, node, type));
        }
    }

    private void AddValue(string name, eColorTransformType type, double value)
    {
        this.AddValue(name, type);
        this._list[this._list.Count - 1].Value = value;
    }

    private XmlElement AddNode(string name)
    {
        XmlElement? node = this._topNode.OwnerDocument.CreateElement("a", name, ExcelPackage.schemaDrawings);
        _ = this._topNode.AppendChild(node);

        return node;
    }

    #endregion

    /// <summary>
    /// Gets the enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<IColorTransformItem> GetEnumerator() => this._list.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this._list.GetEnumerator();

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count => this._list.Count;
}