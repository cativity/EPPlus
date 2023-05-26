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
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style;

/// <summary>
/// Manages colors for a chart style
/// </summary>
public class ExcelChartStyleColorManager : ExcelDrawingColorManager
{
    internal ExcelChartStyleColorManager(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder, Action initMethod = null) : 
        base(nameSpaceManager, topNode, path, schemaNodeOrder, initMethod)  
    {
        if (this._pathNode == null || this._colorNode == null)
        {
            return;
        }

        switch (this._colorNode.LocalName)
        {
            case "styleClr":
                this.ColorType = eDrawingColorType.ChartStyleColor;
                this.StyleColor = new ExcelChartStyleColor(this._nameSpaceManager, this._pathNode.FirstChild);
                break;
        }
    }
    /// <summary>
    /// Sets the style color for a chart style
    /// </summary>
    /// <param name="index">Is index, maps to the style matrix in the theme</param>
    public void SetStyleColor(int index = 0)
    {
        this.SetStyleColor(false, index);
    }
    internal const string NodeName = "a:styleClr";

    /// <summary>
    /// Sets the style color for a chart style
    /// </summary>
    /// <param name="isAuto">Is automatic</param>
    /// <param name="index">Is index, maps to the style matrix in the theme</param>
    public void SetStyleColor(bool isAuto = true, int index = 0)
    {
        this.ColorType = eDrawingColorType.ChartStyleColor;
        this.ResetColors(NodeName);
        this.StyleColor=new ExcelChartStyleColor(this._nameSpaceManager, this._colorNode);

        this.StyleColor.SetValue(isAuto, index);
    }
    /// <summary>
    /// The style color object
    /// </summary>
    public ExcelChartStyleColor StyleColor
    {
        get;
        private set;
    }

    /// <summary>
    /// Reset the color
    /// </summary>
    /// <param name="newNodeName">The new name</param>
    internal protected new void ResetColors(string newNodeName)
    {
        base.ResetColors(newNodeName);
        this.StyleColor = null;
    }
}