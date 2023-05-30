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

using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Filter;

/// <summary>
/// Various filters that are set depending on the filter <c>Type</c>
/// <see cref="Type"/>
/// </summary>
public class ExcelDynamicFilterColumn : ExcelFilterColumn
{
    internal ExcelDynamicFilterColumn(XmlNamespaceManager namespaceManager, XmlNode topNode)
        : base(namespaceManager, topNode)
    {
        DynamicDateFilterMatcher.SetMatchDates(this);
    }

    /// <summary>
    /// Type of filter
    /// </summary>
    public eDynamicFilterType Type { get; set; }

    /// <summary>
    /// The value of the filter. Can be the Average or minimum value depending on the type
    /// </summary>
    public double? Value { get; internal set; }

    /// <summary>
    /// The maximum value for for a daterange, for example ThisMonth
    /// </summary>
    public double? MaxValue { get; internal set; }

    internal override bool Match(object value, string valueText)
    {
        if (this.Type == eDynamicFilterType.AboveAverage)
        {
            return ConvertUtil.GetValueDouble(value) > this.Value;
        }
        else if (this.Type == eDynamicFilterType.BelowAverage)
        {
            return ConvertUtil.GetValueDouble(value) < this.Value;
        }
        else
        {
            DateTime? date = ConvertUtil.GetValueDate(value);

            if (date.HasValue == false)
            {
                return false;
            }

            return DynamicDateFilterMatcher.Match(this, date);
        }
    }

    internal override void Save()
    {
        XmlElement? node = (XmlElement)this.CreateNode("d:dynamicFilter");
        node.RemoveAll();
        //string? type = this.Type.ToEnumString();

        //if (type.Length <= 3)
        //{
        //    type = type.ToUpper(); //For M1, M12, Q1 etc
        //}

        node.SetAttribute("type", GetTypeForXml(this.Type));

        if (this.Value.HasValue)
        {
            node.SetAttribute("val", this.Value.Value.ToString("R15", CultureInfo.InvariantCulture));
        }

        if (this.MaxValue.HasValue)
        {
            node.SetAttribute("maxVal", this.MaxValue.Value.ToString("R15", CultureInfo.InvariantCulture));
        }
    }

    private static string GetTypeForXml(eDynamicFilterType type)
    {
        if (type.ToString().Length > 3)
        {
            return type.ToEnumString();
        }
        else
        {
            return type.ToString();
        }
    }

    internal override void SetFilterValue(ExcelWorksheet worksheet, ExcelAddressBase address)
    {
        if (this.Type == eDynamicFilterType.AboveAverage || this.Type == eDynamicFilterType.BelowAverage)
        {
            this.Value = this.GetAvg(worksheet, address);
            this.MaxValue = null;
        }
        else
        {
            DynamicDateFilterMatcher.SetMatchDates(this);
        }
    }

    private double GetAvg(ExcelWorksheet worksheet, ExcelAddressBase address)
    {
        int count = 0;
        double sum = 0;
        int col = address._fromCol + this.Position;

        for (int row = address._fromRow + 1; row <= address._toRow; row++)
        {
            object? v = worksheet.GetValue(row, col);

            if (ConvertUtil.IsNumericOrDate(v))
            {
                sum += ConvertUtil.GetValueDouble(v);
                count++;
            }
        }

        if (count == 0)
        {
            return 0;
        }
        else
        {
            return sum / count;
        }
    }
}