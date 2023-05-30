using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Utils;
using System;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Handles values with different measurement units. 
/// </summary>
public class ExcelVmlMeasurementUnit
{
    static string[] _suffixes = ((eMeasurementUnits[])Enum.GetValues(typeof(eMeasurementUnits))).Where(x => x != eMeasurementUnits.Pixels)
                                                                                                .Select(x => x.TranslateString())
                                                                                                .ToArray();

    internal ExcelVmlMeasurementUnit(string value = "") => this.SetValue(value);

    /// <summary>
    /// The value of the specified unit.
    /// </summary>
    public double Value { get; set; }

    /// <summary>
    /// The unit of measurement.
    /// </summary>
    public eMeasurementUnits Unit { get; set; }

    internal void SetValue(string value)
    {
        this.Value = GetValue(value);
        this.Unit = GetUnit(value);
    }

    internal string GetValueString() => this.Value.ToString(CultureInfo.InvariantCulture) + this.Unit.TranslateString();

    private static double GetValue(string v)
    {
        if (string.IsNullOrEmpty(v))
        {
            return 0;
        }

        if (_suffixes.Any(x => v.EndsWith(x)))
        {
            return ConvertUtil.GetValueDouble(v.Substring(0, v.Length - 2));
        }

        return ConvertUtil.GetValueDouble(v);
    }

    private static eMeasurementUnits GetUnit(string v)
    {
        foreach (eMeasurementUnits u in Enum.GetValues(typeof(eMeasurementUnits)))
        {
            if (v.EndsWith(u.TranslateString()))
            {
                return u;
            }
        }

        return eMeasurementUnits.Pixels;
    }

    internal double? ToEmu() => VmlConvertUtil.ConvertToEMU(this.Value, this.Unit);
}