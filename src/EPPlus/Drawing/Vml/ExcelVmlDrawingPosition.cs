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
using System.Globalization;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// The position of a VML drawing. Used for comments
/// </summary>
public class ExcelVmlDrawingPosition : XmlHelper
{
    int _startPos;
    internal ExcelVmlDrawingPosition(XmlNamespaceManager ns, XmlNode topNode, int startPos) : 
        base(ns, topNode)
    {
        this._startPos = startPos;
    }
    /// <summary>
    /// Row. Zero based
    /// </summary>
    public int Row
    {
        get
        {
            return this.GetNumber(2);
        }
        set
        {
            this.SetNumber(2, value);
        } 
    }
    /// <summary>
    /// Row offset in pixels. Zero based
    /// </summary>
    public int RowOffset
    {
        get
        {
            return this.GetNumber(3);
        }
        set
        {
            this.SetNumber(3, value);
        }
    }
    /// <summary>
    /// Column. Zero based
    /// </summary>
    public int Column
    {
        get
        {
            return this.GetNumber(0);
        }
        set
        {
            this.SetNumber(0, value);
        }
    }
    /// <summary>
    /// Column offset. Zero based
    /// </summary>
    public int ColumnOffset
    {
        get
        {
            return this.GetNumber(1);
        }
        set
        {
            this.SetNumber(1, value);
        }
    }
    private void SetNumber(int pos, int value)
    {
        string anchor = this.GetXmlNodeString("x:Anchor");
        string[] numbers = anchor.Split(',');
        if (numbers.Length == 8)
        {
            numbers[this._startPos + pos] = value.ToString();
        }
        else
        {
            throw (new Exception("Anchor element is invalid in vmlDrawing"));
        }

        this.SetXmlNodeString("x:Anchor", string.Join(",",numbers));
    }

    private int GetNumber(int pos)
    {
        string anchor = this.GetXmlNodeString("x:Anchor");
        string[] numbers = anchor.Split(',');
        if (numbers.Length == 8)
        {
            if (int.TryParse(numbers[this._startPos + pos], NumberStyles.Number, CultureInfo.InvariantCulture, out int ret))
            {
                return ret;
            }
        }
        throw(new Exception("Anchor element is invalid in vmlDrawing"));
    }
}