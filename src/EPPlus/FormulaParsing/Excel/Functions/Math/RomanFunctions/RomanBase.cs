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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math.RomanFunctions;

internal abstract class RomanBase
{
    protected class RomanNumber
    {
        public RomanNumber(int number, string letter)
        {
            this.Number = number;
            this.Letter = letter;
        }

        public int Number { get; set; }

        public string Letter { get; set; }
    }

    protected readonly RomanNumber One = new RomanNumber(1, "I");
    protected readonly RomanNumber Five = new RomanNumber(5, "V");
    protected readonly RomanNumber Ten = new RomanNumber(10, "X");
    protected readonly RomanNumber Fifty = new RomanNumber(50, "L");
    protected readonly RomanNumber OneHundred = new RomanNumber(100, "C");
    protected readonly RomanNumber FiveHundred = new RomanNumber(500, "D");
    protected readonly RomanNumber Thousand = new RomanNumber(1000, "M");

    internal abstract string Execute(int number);

    protected string GetClassicRomanFormat(int number)
    {
        StringBuilder? roman = new StringBuilder();
        Apply(ref number, this.Thousand, roman);
        Apply(ref number, 900, "CM", roman);
        Apply(ref number, this.FiveHundred, this.OneHundred, roman);
        Apply(ref number, 400, "CD", roman);
        Apply(ref number, this.OneHundred, roman);
        Apply(ref number, 90, "XC", roman);
        Apply(ref number, this.Fifty, this.Ten, roman);
        Apply(ref number, 40, "XL", roman);
        Apply(ref number, this.Ten, roman);
        Apply(ref number, 9, "IX", roman);
        Apply(ref number, this.Five, this.One, roman);
        Apply(ref number, 4, "IV", roman);
        Apply(ref number, this.One, roman);

        return roman.ToString();
    }

    private static void Apply(ref int number, RomanNumber roman, StringBuilder result)
    {
        if (number >= roman.Number)
        {
            int limit = number / roman.Number;

            for (int x = 0; x < limit; x++)
            {
                _ = result.Append(roman.Letter);
                number -= roman.Number;
            }
        }
    }

    private static void Apply(ref int number, RomanNumber roman, RomanNumber lowerRoman, StringBuilder result)
    {
        if (number >= roman.Number)
        {
            _ = result.Append(roman.Letter);
            number -= roman.Number;

            for (int x = 0; x < number / lowerRoman.Number; x++)
            {
                _ = result.Append(lowerRoman.Letter);
                number -= lowerRoman.Number;
            }
        }
    }

    private static void Apply(ref int number, int limit, string letters, StringBuilder result)
    {
        if (number >= limit)
        {
            _ = result.Append(letters);
            number -= limit;
        }
    }

    protected static string HandleType(int type, string roman)
    {
        if (type <= 0)
        {
            return roman;
        }

        // all other types than 0
        roman = roman.Replace("CMXCV", "LMVL");
        roman = roman.Replace("CML", "LM");
        roman = roman.Replace("CDL", "LD");
        roman = roman.Replace("XCV", "VC");
        roman = roman.Replace("XLV", "VL");

        // type 1
        if (type == 1)
        {
            roman = roman.Replace("XLIX", "VLIV");
            roman = roman.Replace("CMXCIX", "LMVLIV");
            roman = roman.Replace("XCIX", "VCIV");
            roman = roman.Replace("CMXC", "LMXL");
            roman = roman.Replace("CDVC", "LDVL");
            roman = roman.Replace("CDXC", "LDXL");
        }

        if (type > 1)
        {
            roman = roman.Replace("XLIX", "IL");
            roman = roman.Replace("XCIX", "IC");
            roman = roman.Replace("CDXC", "XD");
            roman = roman.Replace("CDVC", "XDV");
            roman = roman.Replace("CDIC", "XDIX");
            roman = roman.Replace("LMVL", "XMV");
            roman = roman.Replace("CMIC", "XMIX");
            roman = roman.Replace("CMXC", "XM");
        }

        if (type > 2)
        {
            roman = roman.Replace("XDV", "VD");
            roman = roman.Replace("XDIX", "VDIV");
            roman = roman.Replace("XMV", "VM");
            roman = roman.Replace("XMIX", "VMIV");
        }

        if (type == 4)
        {
            roman = roman.Replace("VDIV", "ID");
            roman = roman.Replace("VMIV", "IM");
        }

        return roman;
    }
}