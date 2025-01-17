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
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style;

/// <summary>
/// Represents a color style of a chart.
/// </summary>
public class ExcelChartColorsManager : XmlHelper
{
    internal ExcelChartColorsManager(XmlNamespaceManager nameSpaceManager, XmlElement topNode)
        : base(nameSpaceManager, topNode)
    {
        List<ExcelDrawingColorManager>? colors = new List<ExcelDrawingColorManager>();
        List<ExcelColorTransformCollection>? variations = new List<ExcelColorTransformCollection>();

        foreach (XmlElement c in topNode.ChildNodes)
        {
            if (c.LocalName == "variation")
            {
                variations.Add(new ExcelColorTransformCollection(this.NameSpaceManager, c));
            }
            else
            {
                colors.Add(new ExcelDrawingColorManager(this.NameSpaceManager, c, "", new string[0]));
            }
        }

        this.Colors = new ReadOnlyCollection<ExcelDrawingColorManager>(colors);
        this.Variations = new ReadOnlyCollection<ExcelColorTransformCollection>(variations);
    }

    /// <summary>
    /// The method to use to calculate the colors
    /// </summary>
    /// <remarks>AcrossLinear is not implemented yet, and will use WithinLinear</remarks>
    public eChartColorStyleMethod Method
    {
        get => this.GetXmlNodeString("@meth").ToEnum(eChartColorStyleMethod.Cycle);
        set => this.SetXmlNodeString("@meth", value.ToEnumString());
    }

    /// <summary>
    /// The colors to use for the calculation
    /// </summary>
    public ReadOnlyCollection<ExcelDrawingColorManager> Colors { get; }

    /// <summary>
    /// The variations to use for the calculation
    /// </summary>
    public ReadOnlyCollection<ExcelColorTransformCollection> Variations { get; }

    internal void Transform(ExcelDrawingColorManager color, int colorIndex, int numberOfItems)
    {
        ExcelDrawingColorManager? newColor = this.GetColor(colorIndex, numberOfItems);
        ExcelColorTransformCollection? variation = this.GetVariation(colorIndex, numberOfItems);
        color.ApplyNewColor(newColor, variation);
    }

    private ExcelDrawingColorManager GetColor(int colorIndex, int numberOfItems)
    {
        switch (this.Method)
        {
            case eChartColorStyleMethod.Cycle:
                int ix = colorIndex % this.Colors.Count;

                return this.Colors[ix];

            default:
                //TODO add support for other types.
                ix = colorIndex % this.Colors.Count;

                return this.Colors[ix];
        }
    }

    private ExcelColorTransformCollection GetVariation(int colorIndex, int numberOfItems)
    {
        switch (this.Method)
        {
            case eChartColorStyleMethod.AcrossLinear:
            case eChartColorStyleMethod.WithinLinear:
                return GetLinearVariation(colorIndex, numberOfItems, false);

            case eChartColorStyleMethod.AcrossLinearReversed:
            case eChartColorStyleMethod.WithinLinearReversed:
                return GetLinearVariation(colorIndex, numberOfItems, true);

            //eChartColorStyleMethod.Cycle
            default:
                int div = colorIndex - (colorIndex % this.Colors.Count);

                if (div == 0)
                {
                    return this.Variations[0];
                }
                else
                {
                    int ix = div / this.Colors.Count;
                    ix %= this.Variations.Count;

                    return this.Variations[ix];
                }
        }
    }

    private static ExcelColorTransformCollection GetLinearVariation(int colorIndex, int numberOfItems, bool isReversed)
    {
        ExcelColorTransformCollection? ret = new ExcelColorTransformCollection();

        if (numberOfItems <= 1)
        {
            return ret;
        }

        double split = (numberOfItems - 1) / 2D;

        if (colorIndex == split)
        {
            return ret;
        }
        else
        {
            int span = GetVariationStart(numberOfItems / 2D);
            double diff = (double)span / split;
            bool isTint;
            int v;

            if (colorIndex > split)
            {
                v = (int)(100 - (diff * -(split - colorIndex)));
                isTint = !isReversed;
            }
            else
            {
                v = (int)(100 - (diff * (split - colorIndex)));
                isTint = isReversed;
            }

            if (isTint)
            {
                ret.AddTint(v);
            }
            else
            {
                ret.AddShade(v);
            }
        }

        return ret;
    }

    private static int GetVariationStart(double split)
    {
        int diff = 24;
        int ret = 0;

        while (split > 0)
        {
            ret += diff;

            if (ret >= 80)
            {
                break;
            }

            split--;

            if (split < 1)
            {
                ret += (int)(diff * split) >> 1;

                break;
            }

            if (diff > 1)
            {
                diff >>= 1;
            }
        }

        return ret;
    }
}