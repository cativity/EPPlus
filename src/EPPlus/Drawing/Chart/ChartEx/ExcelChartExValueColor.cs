/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/

using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// Individual color settings for a region map charts series colors
/// </summary>
public class ExcelChartExValueColor : XmlHelper
{
    string _prefix;
    string _positionPath;

    internal ExcelChartExValueColor(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string prefix)
        : base(nameSpaceManager, topNode)
    {
        this.SchemaNodeOrder = schemaNodeOrder;
        this._prefix = prefix;
        this._positionPath = $"cx:valueColorPositions/cx:{prefix}Position";
    }

    ExcelDrawingColorManager _color;

    /// <summary>
    /// The color
    /// </summary>
    public ExcelDrawingColorManager Color
    {
        get
        {
            return this._color ??= new ExcelDrawingColorManager(this.NameSpaceManager,
                                                                this.TopNode,
                                                                $"cx:valueColors/cx:{this._prefix}Color",
                                                                this.SchemaNodeOrder);
        }
    }

    /// <summary>
    /// The color variation type.
    /// </summary>
    public eColorValuePositionType ValueType
    {
        get
        {
            if (this.ExistsNode($"{this._positionPath}/cx:number"))
            {
                return eColorValuePositionType.Number;
            }
            else if (this.ExistsNode($"{this._positionPath}/cx:percent"))
            {
                return eColorValuePositionType.Percent;
            }
            else
            {
                return eColorValuePositionType.Extreme;
            }
        }
        set
        {
            if (this.ValueType != value)
            {
                this.ClearChildren(this._positionPath);

                switch (value)
                {
                    case eColorValuePositionType.Extreme:
                        _ = this.CreateNode($"{this._positionPath}/cx:extremeValue");

                        break;

                    case eColorValuePositionType.Percent:
                        this.SetXmlNodeString($"{this._positionPath}/cx:percent/@val", "0");

                        break;

                    default:
                        this.SetXmlNodeString($"{this._positionPath}/cx:number/@val", "0");

                        break;
                }
            }
        }
    }

    /// <summary>
    /// The color variation value.
    /// </summary>
    public double PositionValue
    {
        get
        {
            eColorValuePositionType t = this.ValueType;

            if (t == eColorValuePositionType.Extreme)
            {
                return 0;
            }
            else if (this.ValueType == eColorValuePositionType.Number)
            {
                return this.GetXmlNodeDouble($"{this._positionPath}/cx:number/@val");
            }
            else
            {
                return this.GetXmlNodeDoubleNull($"{this._positionPath}/cx:percent/@val") ?? 0;
            }
        }
        set
        {
            eColorValuePositionType t = this.ValueType;

            if (t == eColorValuePositionType.Extreme)
            {
                throw new InvalidOperationException("Can't set PositionValue when ValueType is Extreme");
            }
            else if (t == eColorValuePositionType.Number)
            {
                this.SetXmlNodeString($"{this._positionPath}/cx:number/@val", value.ToString(CultureInfo.InvariantCulture));
            }
            else if (t == eColorValuePositionType.Percent)
            {
                if (value < 0 || value > 100)
                {
                    throw new InvalidOperationException("PositionValue out of range. Percantage range is from 0 to 100");
                }

                this.SetXmlNodeDouble($"{this._positionPath}/cx:percent/@val", value);
            }
        }
    }
}