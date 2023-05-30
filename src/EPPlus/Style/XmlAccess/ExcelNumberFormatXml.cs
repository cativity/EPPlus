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
using System.Text;
using System.Xml;
using System.Globalization;
using System.Text.RegularExpressions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Style.XmlAccess;

/// <summary>
/// Xml access class for number formats
/// </summary>
public sealed class ExcelNumberFormatXml : StyleXmlHelper
{
    internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager)
        : base(nameSpaceManager)
    {
    }

    internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager, bool buildIn)
        : base(nameSpaceManager) =>
        this.BuildIn = buildIn;

    internal ExcelNumberFormatXml(XmlNamespaceManager nsm, XmlNode topNode)
        : base(nsm, topNode)
    {
        this._numFmtId = this.GetXmlNodeInt("@numFmtId");
        this._format = this.GetXmlNodeString("@formatCode");
    }

    /// <summary>
    /// If the number format is build in
    /// </summary>
    public bool BuildIn { get; private set; }

    int _numFmtId;

    /// <summary>
    /// Id for number format
    /// 
    /// Build in ID's
    /// 
    /// 0   General 
    /// 1   0 
    /// 2   0.00 
    /// 3   #,##0 
    /// 4   #,##0.00 
    /// 9   0% 
    /// 10  0.00% 
    /// 11  0.00E+00 
    /// 12  # ?/? 
    /// 13  # ??/?? 
    /// 14  mm-dd-yy 
    /// 15  d-mmm-yy 
    /// 16  d-mmm 
    /// 17  mmm-yy 
    /// 18  h:mm AM/PM 
    /// 19  h:mm:ss AM/PM 
    /// 20  h:mm 
    /// 21  h:mm:ss 
    /// 22  m/d/yy h:mm 
    /// 37  #,##0;(#,##0) 
    /// 38  #,##0;[Red] (#,##0) 
    /// 39  #,##0.00;(#,##0.00) 
    /// 40  #,##0.00;[Red] (#,##0.00) 
    /// 45  mm:ss 
    /// 46  [h]:mm:ss 
    /// 47  mmss.0 
    /// 48  ##0.0E+0 
    /// 49  @
    /// </summary>            
    public int NumFmtId
    {
        get => this._numFmtId;
        set => this._numFmtId = value;
    }

    internal override string Id => this._format;

    const string fmtPath = "@formatCode";
    string _format = string.Empty;

    /// <summary>
    /// The numberformat string
    /// </summary>
    public string Format
    {
        get => this._format;
        set
        {
            this._numFmtId = ExcelNumberFormat.GetFromBuildIdFromFormat(value);
            this._format = value;
        }
    }

    internal static string GetNewID(int NumFmtId, string Format)
    {
        if (NumFmtId < 0)
        {
            NumFmtId = ExcelNumberFormat.GetFromBuildIdFromFormat(Format);
        }

        return NumFmtId.ToString();
    }

    internal static void AddBuildIn(XmlNamespaceManager NameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats)
    {
        _ = NumberFormats.Add("General", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 0, Format = "General" });
        _ = NumberFormats.Add("0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 1, Format = "0" });
        _ = NumberFormats.Add("0.00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 2, Format = "0.00" });
        _ = NumberFormats.Add("#,##0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 3, Format = "#,##0" });
        _ = NumberFormats.Add("#,##0.00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 4, Format = "#,##0.00" });
        _ = NumberFormats.Add("0%", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 9, Format = "0%" });
        _ = NumberFormats.Add("0.00%", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 10, Format = "0.00%" });
        _ = NumberFormats.Add("0.00E+00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 11, Format = "0.00E+00" });
        _ = NumberFormats.Add("# ?/?", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 12, Format = "# ?/?" });
        _ = NumberFormats.Add("# ??/??", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 13, Format = "# ??/??" });
        _ = NumberFormats.Add("mm-dd-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 14, Format = "mm-dd-yy" });
        _ = NumberFormats.Add("d-mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 15, Format = "d-mmm-yy" });
        _ = NumberFormats.Add("d-mmm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 16, Format = "d-mmm" });
        _ = NumberFormats.Add("mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 17, Format = "mmm-yy" });
        _ = NumberFormats.Add("h:mm AM/PM", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 18, Format = "h:mm AM/PM" });
        _ = NumberFormats.Add("h:mm:ss AM/PM", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 19, Format = "h:mm:ss AM/PM" });
        _ = NumberFormats.Add("h:mm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 20, Format = "h:mm" });
        _ = NumberFormats.Add("h:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 21, Format = "h:mm:ss" });
        _ = NumberFormats.Add("m/d/yy h:mm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 22, Format = "m/d/yy h:mm" });
        _ = NumberFormats.Add("#,##0 ;(#,##0)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 37, Format = "#,##0 ;(#,##0)" });
        _ = NumberFormats.Add("#,##0 ;[Red](#,##0)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 38, Format = "#,##0 ;[Red](#,##0)" });
        _ = NumberFormats.Add("#,##0.00;(#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 39, Format = "#,##0.00;(#,##0.00)" });
        _ = NumberFormats.Add("#,##0.00;[Red](#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 40, Format = "#,##0.00;[Red](#,##0.00)" });
        _ = NumberFormats.Add("mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 45, Format = "mm:ss" });
        _ = NumberFormats.Add("[h]:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 46, Format = "[h]:mm:ss" });
        _ = NumberFormats.Add("mmss.0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 47, Format = "mmss.0" });
        _ = NumberFormats.Add("##0.0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 48, Format = "##0.0" });
        _ = NumberFormats.Add("@", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 49, Format = "@" });

        NumberFormats.NextId = 164; //Start for custom formats.
    }

    internal override XmlNode CreateXmlNode(XmlNode topNode)
    {
        this.TopNode = topNode;
        this.SetXmlNodeString("@numFmtId", this.NumFmtId.ToString());
        this.SetXmlNodeString("@formatCode", this.Format);

        return this.TopNode;
    }

    internal enum eFormatType
    {
        Unknown = 0,
        Number = 1,
        DateTime = 2,
    }

    ExcelFormatTranslator _translator;

    internal ExcelFormatTranslator FormatTranslator => this._translator ??= new ExcelFormatTranslator(this.Format, this.NumFmtId);

    #region Excel --> .Net Format

    internal class ExcelFormatTranslator
    {
        internal enum eSystemDateFormat
        {
            None,
            SystemLongDate,
            SystemLongTime,
            Conditional,
            SystemShortDate,
        }

        internal class FormatPart
        {
            internal string NetFormat { get; set; }

            internal string NetFormatForWidth { get; set; }

            internal string FractionFormat { get; set; }

            internal eSystemDateFormat SpecialDateFormat { get; set; }

            internal bool ContainsTextPlaceholder { get; set; }

            internal void SetFormat(string format, bool containsAmPm, bool forColWidth)
            {
                if (containsAmPm)
                {
                    format += "tt";
                }

                if (forColWidth)
                {
                    this.NetFormatForWidth = format;
                }
                else
                {
                    this.NetFormat = format;
                }
            }
        }

        internal ExcelFormatTranslator(string format, int numFmtID)
        {
            FormatPart? f = new FormatPart();
            this.Formats.Add(f);

            if (numFmtID == 14)
            {
                f.NetFormat = f.NetFormatForWidth = "";
                this.DataType = eFormatType.DateTime;
                f.SpecialDateFormat = eSystemDateFormat.SystemShortDate;
            }
            else if (format.Equals("general", StringComparison.OrdinalIgnoreCase))
            {
                f.NetFormat = f.NetFormatForWidth = "0.#########";
                this.DataType = eFormatType.Number;
            }
            else
            {
                this.ToNetFormat(format, false);
                this.ToNetFormat(format, true);
            }
        }

        internal List<FormatPart> Formats { get; private set; } = new List<FormatPart>();

        CultureInfo _ci;

        internal CultureInfo Culture
        {
            get => this._ci ?? CultureInfo.CurrentCulture;
            set => this._ci = value;
        }

        internal bool HasCulture => this._ci != null;

        internal eFormatType DataType { get; private set; }

        private void ToNetFormat(string ExcelFormat, bool forColWidth)
        {
            this.DataType = eFormatType.Unknown;
            bool isText = false;
            bool isBracket = false;
            string bracketText = "";
            bool prevBslsh = false;
            bool useMinute = false;
            bool prevUnderScore = false;
            bool ignoreNext = false;
            bool containsAmPm = ExcelFormat.Contains("AM/PM");
            List<int> lstDec = new List<int>();
            StringBuilder sb = new StringBuilder();
            this.Culture = null;
            int secCount = 0;
            FormatPart? f = this.Formats[0];

            if (containsAmPm)
            {
                ExcelFormat = Regex.Replace(ExcelFormat, "AM/PM", "");
                this.DataType = eFormatType.DateTime;
            }

            for (int pos = 0; pos < ExcelFormat.Length; pos++)
            {
                char c = ExcelFormat[pos];

                if (c == '"')
                {
                    isText = !isText;
                    _ = sb.Append(c);
                }
                else
                {
                    if (ignoreNext)
                    {
                        ignoreNext = false;

                        continue;
                    }
                    else if (isText && !isBracket)
                    {
                        _ = sb.Append(c);
                    }
                    else if (isBracket)
                    {
                        if (c == ']')
                        {
                            isBracket = false;

                            if (bracketText[0] == '$') //Local Info
                            {
                                string[] li = Regex.Split(bracketText, "-");

                                if (li[0].Length > 1)
                                {
                                    _ = sb.Append("\"" + li[0].Substring(1, li[0].Length - 1) + "\""); //Currency symbol
                                }

                                if (li.Length > 1)
                                {
                                    if (li[1].Equals("f800", StringComparison.OrdinalIgnoreCase))
                                    {
                                        f.SpecialDateFormat = eSystemDateFormat.SystemLongDate;
                                    }
                                    else if (li[1].Equals("f400", StringComparison.OrdinalIgnoreCase))
                                    {
                                        f.SpecialDateFormat = eSystemDateFormat.SystemLongTime;
                                    }
                                    else if (int.TryParse(li[1], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int num))
                                    {
                                        try
                                        {
                                            this.Culture = CultureInfo.GetCultureInfo(num & 0xFFFF);
                                        }
                                        catch
                                        {
                                            this.Culture = null;
                                        }
                                    }
                                    else //Excel saves in hex, but seems to support having culture codes as well.
                                    {
                                        try
                                        {
                                            this.Culture = CultureInfo.GetCultureInfo(li[1]);
                                        }
                                        catch
                                        {
                                            this.Culture = null;
                                        }
                                    }
                                }
                            }
                            else if (bracketText.StartsWith("<") || bracketText.StartsWith(">") || bracketText.StartsWith("=")) //Conditional
                            {
                                f.SpecialDateFormat = eSystemDateFormat.Conditional;
                            }
                            else
                            {
                                _ = sb.Append(bracketText);
                                f.SpecialDateFormat = eSystemDateFormat.Conditional;
                            }
                        }
                        else
                        {
                            bracketText += c;
                        }
                    }
                    else if (prevUnderScore)
                    {
                        if (forColWidth)
                        {
                            _ = sb.AppendFormat("\"{0}\"", c);
                        }

                        prevUnderScore = false;
                    }
                    else
                    {
                        if (c == ';') //We use first part (for positive only at this stage)
                        {
                            secCount++;
                            f.SetFormat(sb.ToString(), containsAmPm, forColWidth);

                            if (secCount < this.Formats.Count)
                            {
                                f = this.Formats[secCount];
                            }
                            else
                            {
                                f = new FormatPart();
                                this.Formats.Add(f);
                            }

                            sb = new StringBuilder();
                        }
                        else
                        {
                            char clc = c.ToString().ToLower(CultureInfo.InvariantCulture)[0]; //Lowercase character

                            //Set the datetype
                            if (this.DataType == eFormatType.Unknown)
                            {
                                if (c == '0' || c == '#' || c == '.')
                                {
                                    this.DataType = eFormatType.Number;
                                }
                                else if (clc == 'y' || clc == 'm' || clc == 'd' || clc == 'h' || clc == 'm' || clc == 's')
                                {
                                    this.DataType = eFormatType.DateTime;
                                }
                            }

                            if (prevBslsh)
                            {
                                if (c == '.' || c == ',')
                                {
                                    _ = sb.Append('\\');
                                }

                                _ = sb.Append(c);
                                prevBslsh = false;
                            }
                            else if (c == '[')
                            {
                                bracketText = "";
                                isBracket = true;
                            }
                            else if (c == '\\')
                            {
                                prevBslsh = true;
                            }
                            else if (c == '0' || c == '#' || c == '.' || c == ',' || c == '%' || clc == 'd' || clc == 's')
                            {
                                _ = sb.Append(c);

                                if (c == '.')
                                {
                                    lstDec.Add(sb.Length - 1);
                                }
                            }
                            else if (clc == 'h')
                            {
                                if (containsAmPm)
                                {
                                    _ = sb.Append('h');
                                }
                                else
                                {
                                    _ = sb.Append('H');
                                }

                                useMinute = true;
                            }
                            else if (clc == 'm')
                            {
                                if (useMinute)
                                {
                                    _ = sb.Append('m');
                                }
                                else
                                {
                                    _ = sb.Append('M');
                                }
                            }
                            else if (c == '_') //Skip next but use for alignment
                            {
                                prevUnderScore = true;
                            }
                            else if (c == '?')
                            {
                                _ = sb.Append(' ');
                            }
                            else if (c == '/')
                            {
                                if (this.DataType == eFormatType.Number)
                                {
                                    int startPos = pos - 1;

                                    while (startPos >= 0 && (ExcelFormat[startPos] == '?' || ExcelFormat[startPos] == '#' || ExcelFormat[startPos] == '0'))
                                    {
                                        startPos--;
                                    }

                                    if (startPos > 0) //RemovePart
                                    {
                                        _ = sb.Remove(sb.Length - (pos - startPos - 1), pos - startPos - 1);
                                    }

                                    int endPos = pos + 1;

                                    while (endPos < ExcelFormat.Length
                                           && (ExcelFormat[endPos] == '?'
                                               || ExcelFormat[endPos] == '#'
                                               || (ExcelFormat[endPos] >= '0' && ExcelFormat[endPos] <= '9')))
                                    {
                                        endPos++;
                                    }

                                    pos = endPos;

                                    if (f.FractionFormat != "")
                                    {
                                        f.FractionFormat = ExcelFormat.Substring(startPos + 1, endPos - startPos - 1);
                                    }

                                    _ = sb.Append('?'); //Will be replaced later on by the fraction
                                }
                                else
                                {
                                    _ = sb.Append('/');
                                }
                            }
                            else if (c == '*')
                            {
                                //repeat char--> ignore
                                ignoreNext = true;
                            }
                            else if (c == '@')
                            {
                                _ = sb.Append("{0}");
                                f.ContainsTextPlaceholder = true;
                            }
                            else
                            {
                                _ = sb.Append(c);
                            }
                        }
                    }
                }
            }

            //Add qoutes
            if (this.DataType == eFormatType.DateTime)
            {
                SetDecimal(lstDec, sb); //Remove?
            }

            //if (format == "")
            //    format = sb.ToString();
            //else
            //    text = sb.ToString();

            // AM/PM format
            f.SetFormat(sb.ToString(), containsAmPm, forColWidth);
        }

        private static void SetDecimal(List<int> lstDec, StringBuilder sb)
        {
            if (lstDec.Count > 1)
            {
                for (int i = lstDec.Count - 1; i >= 0; i--)
                {
                    _ = sb.Insert(lstDec[i] + 1, '\'');
                    _ = sb.Insert(lstDec[i], '\'');
                }
            }
        }

        internal static string FormatFraction(double d, FormatPart f)
        {
            int numerator,
                denomerator;

            int intPart = (int)d;

            string[] fmt = f.FractionFormat.Split('/');

            if (!int.TryParse(fmt[1], out int fixedDenominator))
            {
                fixedDenominator = 0;
            }

            if (d == 0 || double.IsNaN(d))
            {
                if (fmt[0].Trim() == "" && fmt[1].Trim() == "")
                {
                    return new string(' ', f.FractionFormat.Length);
                }
                else
                {
                    return 0.ToString(fmt[0]) + "/" + 1.ToString(fmt[0]);
                }
            }

            int maxDigits = fmt[1].Length;
            string sign = d < 0 ? "-" : "";

            if (fixedDenominator == 0)
            {
                List<double> numerators = new List<double>() { 1, 0 };
                List<double> denominators = new List<double>() { 0, 1 };

                if (maxDigits < 1 && maxDigits > 12)
                {
                    throw new ArgumentException("Number of digits out of range (1-12)");
                }

                int maxNum = 0;

                for (int i = 0; i < maxDigits; i++)
                {
                    maxNum += 9 * (int)Math.Pow((double)10, (double)i);
                }

                double divRes = 1 / ((double)Math.Abs(d) - intPart);
                double prevResult = double.NaN;

                int listPos = 2,
                    index = 1;

                while (true)
                {
                    index++;
                    double intDivRes = Math.Floor(divRes);
                    numerators.Add((intDivRes * numerators[index - 1]) + numerators[index - 2]);

                    if (numerators[index] > maxNum)
                    {
                        break;
                    }

                    denominators.Add((intDivRes * denominators[index - 1]) + denominators[index - 2]);

                    double result = numerators[index] / denominators[index];

                    if (denominators[index] > maxNum)
                    {
                        break;
                    }

                    listPos = index;

                    if (result == prevResult)
                    {
                        break;
                    }

                    if (result == d)
                    {
                        break;
                    }

                    prevResult = result;

                    divRes = 1 / (divRes - intDivRes); //Rest
                }

                numerator = (int)numerators[listPos];
                denomerator = (int)denominators[listPos];
            }
            else
            {
                numerator = (int)Math.Round((d - intPart) / (1D / fixedDenominator), 0);
                denomerator = fixedDenominator;
            }

            if (numerator == denomerator || numerator == 0)
            {
                if (numerator == denomerator)
                {
                    intPart++;
                }

                return sign + intPart.ToString(f.NetFormat).Replace("?", new string(' ', f.FractionFormat.Length));
            }
            else if (intPart == 0)
            {
                return sign + FmtInt(numerator, fmt[0]) + "/" + FmtInt(denomerator, fmt[1]);
            }
            else
            {
                return sign + intPart.ToString(f.NetFormat).Replace("?", FmtInt(numerator, fmt[0]) + "/" + FmtInt(denomerator, fmt[1]));
            }
        }

        private static string FmtInt(double value, string format)
        {
            string v = value.ToString("#");
            string pad = "";

            if (v.Length < format.Length)
            {
                for (int i = format.Length - v.Length - 1; i >= 0; i--)
                {
                    if (format[i] == '?')
                    {
                        pad += " ";
                    }
                    else if (format[i] == ' ')
                    {
                        pad += "0";
                    }
                }
            }

            return pad + v;
        }

        internal FormatPart GetFormatPart(object value)
        {
            if (this.Formats.Count > 1)
            {
                if (ConvertUtil.IsNumericOrDate(value))
                {
                    double d = ConvertUtil.GetValueDouble(value);

                    if (d < 0D && this.Formats.Count > 1)
                    {
                        return this.Formats[1];
                    }
                    else if (d == 0D && this.Formats.Count > 2)
                    {
                        return this.Formats[2];
                    }
                    else
                    {
                        return this.Formats[0];
                    }
                }
                else
                {
                    if (this.Formats.Count > 3)
                    {
                        return this.Formats[3];
                    }
                    else
                    {
                        return this.Formats[0];
                    }
                }
            }
            else
            {
                return this.Formats[0];
            }
        }
    }

    #endregion
}