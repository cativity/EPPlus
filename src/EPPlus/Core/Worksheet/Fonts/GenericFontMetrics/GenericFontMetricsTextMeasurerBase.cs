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
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal abstract class GenericFontMetricsTextMeasurerBase
    {
        private FontScaleFactors _fontScaleFactors = new FontScaleFactors();
        private static Dictionary<uint, SerializedFontMetrics> _fonts;
        private static object _syncRoot = new object();

        public GenericFontMetricsTextMeasurerBase()
        {
            Initialize();
        }

        private static void Initialize()
        {
            lock (_syncRoot)
            {
                if (_fonts == null)
                {
                    _fonts = GenericFontMetricsLoader.LoadFontMetrics();
                }
            }
        }

        protected internal static bool IsValidFont(uint fontKey)
        {
            return _fonts.ContainsKey(fontKey);
        }

        internal protected TextMeasurement MeasureTextInternal(string text, uint fontKey, MeasurementFontStyles style, float size)
        {
            SerializedFontMetrics? sFont = _fonts[fontKey];
            float width = 0f;
            float widthEA = 0f;
            char[]? chars = text.ToCharArray();
            for (int x = 0; x < chars.Length; x++)
            {
                SerializedFontMetrics? fnt = sFont;
                char c = chars[x];
                // if east asian char use default regardless of actual font.
                if (IsEastAsianChar(c))
                {
                    widthEA += GetEastAsianCharWidth(c, style);
                }
                else
                {
                    if (sFont.CharMetrics.ContainsKey(c))
                    {
                        float fw = fnt.ClassWidths[sFont.CharMetrics[c]];
                        if (Char.IsDigit(c))
                        {
                            fw *= FontScaleFactors.DigitsScalingFactor;
                        }

                        width += fw;
                    }
                    else
                    {
                        width += sFont.ClassWidths[fnt.DefaultWidthClass];
                    }
                }

            }
            width *= size;
            widthEA *= size;
            float sf = _fontScaleFactors.GetScaleFactor(fontKey, width);
            width *= sf;
            width += widthEA;
            float height = sFont.LineHeight1em * size;
            return new TextMeasurement(width, height);
        }



        internal static uint GetKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            ushort k1 = (ushort)family;
            ushort k2 = (ushort)subFamily;
            return (uint)((k1 << 16) | ((k2) & 0xffff));
        }

        internal static uint GetKey(string fontFamily, MeasurementFontStyles fontStyle)
        {
            string? enumName = fontFamily.Replace(" ", string.Empty);
            Array? values = Enum.GetValues(typeof(FontMetricsFamilies));
            bool supported = false;
            foreach (object? enumVal in values)
            {
                if (enumVal.ToString() == enumName)
                {
                    supported = true;
                    break;
                }
            }
            if (!supported)
            {
                return uint.MaxValue;
            }

            FontMetricsFamilies family = (FontMetricsFamilies)Enum.Parse(typeof(FontMetricsFamilies), enumName);
            FontSubFamilies subFamily = FontSubFamilies.Regular;
            switch (fontStyle)
            {
                case MeasurementFontStyles.Bold:
                    subFamily = FontSubFamilies.Bold;
                    break;
                case MeasurementFontStyles.Italic:
                    subFamily = FontSubFamilies.Italic;
                    break;
                case MeasurementFontStyles.Italic | MeasurementFontStyles.Bold:
                    subFamily = FontSubFamilies.BoldItalic;
                    break;
                default:
                    break;
            }
            return GetKey(family, subFamily);
        }

        private static float GetEastAsianCharWidth(int cc, MeasurementFontStyles style)
        {
            float emWidth = (cc >= 65377 && cc <= 65439) ? 0.5f : 1f;
            if ((style & MeasurementFontStyles.Bold) != 0)
            {
                emWidth *= 1.05f;
            }
            return emWidth * (96F / 72F) * FontScaleFactors.JapaneseKanjiDefaultScalingFactor;
        }

        private static bool IsEastAsianChar(char c)
        {
            int cc = (int)c;

            return UniCodeRange.JapaneseKanji.Any(x => x.IsInRange(cc));
        }

    }
}
