/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/

using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.SystemDrawing.Text;
using System;

namespace OfficeOpenXml;

/// <summary>
/// This class contains settings for text measurement.
/// </summary>
public class ExcelTextSettings
{
    internal ExcelTextSettings()
    {
        if (Environment.OSVersion.Platform == PlatformID.Unix || Environment.OSVersion.Platform == PlatformID.MacOSX)
        {
            this.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();

            try
            {
                this.FallbackTextMeasurer = new SystemDrawingTextMeasurer();
            }
            catch
            {
                this.FallbackTextMeasurer = null;
            }
        }
        else
        {
            try
            {
                SystemDrawingTextMeasurer? m = new SystemDrawingTextMeasurer();

                if (m.ValidForEnvironment())
                {
                    this.PrimaryTextMeasurer = m;
                    this.FallbackTextMeasurer = new GenericFontMetricsTextMeasurer();
                }
                else
                {
                    this.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
                }
            }
            catch
            {
                this.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
            }
        }

        this.AutofitScaleFactor = 1f;
    }

    /// <summary>
    /// This is the primary text measurer
    /// </summary>
    public ITextMeasurer PrimaryTextMeasurer { get; set; }

    /// <summary>
    /// If the primary text measurer fails to measure the text, this one will be used.
    /// </summary>
    public ITextMeasurer FallbackTextMeasurer { get; set; }

    /// <summary>
    /// All measurements of texts will be multiplied with this value. Default is 1.
    /// </summary>
    public float AutofitScaleFactor { get; set; }

    /// <summary>
    /// Returns an instance of the internal generic text measurer
    /// </summary>
    public static ITextMeasurer GenericTextMeasurer => new GenericFontMetricsTextMeasurer();

    /// <summary>
    /// Measures a text with default settings when there is no other option left...
    /// </summary>
    internal DefaultTextMeasurer DefaultTextMeasurer { get; set; }
}