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
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Coloring;

/// <summary>
/// Different types of transformation performed on a color 
/// </summary>
public class ExcelColorTransformItem : XmlHelper, IColorTransformItem, ISource
{
    internal ExcelColorTransformItem(XmlNamespaceManager nsm, XmlNode topNode, eColorTransformType type)
        : base(nsm, topNode)
    {
        this.Type = type;
        this.DataType = GetDataType(type);
    }

    private static eColorTransformDataType GetDataType(eColorTransformType type)
    {
        switch (type)
        {
            case eColorTransformType.Alpha:
            case eColorTransformType.AlphaMod:
            case eColorTransformType.Tint:
            case eColorTransformType.Shade:
            case eColorTransformType.HueMod:
            case eColorTransformType.Lum:
            case eColorTransformType.Sat:
                return eColorTransformDataType.FixedPositivePercentage;

            case eColorTransformType.AlphaOff:
                return eColorTransformDataType.FixedPercentage;

            case eColorTransformType.HueOff:
                return eColorTransformDataType.Angle;

            case eColorTransformType.Hue:
                return eColorTransformDataType.FixedAngle90;

            case eColorTransformType.Inv:
            case eColorTransformType.Comp:
            case eColorTransformType.Gray:
            case eColorTransformType.Gamma:
            case eColorTransformType.InvGamma:
                return eColorTransformDataType.Boolean;

            default:
                return eColorTransformDataType.Percentage;
        }
    }

    /// <summary>
    /// The type of transformation
    /// </summary>
    public eColorTransformType Type { get; private set; }

    /// <summary>
    /// Datatype for color transformation
    /// </summary>
    public eColorTransformDataType DataType { get; private set; }

    /// <summary>
    /// The value of the color tranformation
    /// </summary>
    public double Value
    {
        get
        {
            switch (this.DataType)
            {
                case eColorTransformDataType.Percentage:
                case eColorTransformDataType.PositivePercentage:
                case eColorTransformDataType.FixedPercentage:
                case eColorTransformDataType.FixedPositivePercentage:
                    return this.GetXmlNodePercentage("@val") ?? 0;

                case eColorTransformDataType.Angle:
                case eColorTransformDataType.FixedAngle90:
                    return this.GetXmlNodeAngel("@val");

                default:
                    return 1; //Boolean
            }
        }
        set
        {
            if (this.DataType == eColorTransformDataType.Boolean)
            {
                throw new ArgumentException("Value",
                                            "Value property don't apply to transformations with datatype Boolean. Please add(true)/remove(false) this item to change it's state");
            }

            if (this.DataType == eColorTransformDataType.Percentage)
            {
                this.SetXmlNodePercentage("@val", value, true, int.MaxValue / 1000);
            }
            else if (this.DataType == eColorTransformDataType.PositivePercentage)
            {
                this.SetXmlNodePercentage("@val", value, false, int.MaxValue / 1000);
            }
            else if (this.DataType == eColorTransformDataType.FixedPercentage)
            {
                this.SetXmlNodePercentage("@val", value);
            }
            else if (this.DataType == eColorTransformDataType.FixedPositivePercentage)
            {
                this.SetXmlNodePercentage("@val", value, false);
            }
            else if (this.DataType == eColorTransformDataType.Angle)
            {
                this.SetXmlNodeAngel("@val", value, this.Type.ToString(), int.MinValue / 60000, int.MaxValue / 60000);
            }
            else if (this.DataType == eColorTransformDataType.FixedAngle90)
            {
                this.SetXmlNodeAngel("@val", value, this.Type.ToString(), -90, 90);
            }
        }
    }

    bool ISource._fromStyleTemplate { get; set; } = false;

    /// <summary>
    /// Converts the object to a string
    /// </summary>
    /// <returns>The type</returns>
    public override string ToString()
    {
        return this.Type.ToString();
    }
}