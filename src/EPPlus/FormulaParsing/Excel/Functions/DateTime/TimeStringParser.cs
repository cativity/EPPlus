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
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

internal class TimeStringParser
{
    private const string RegEx24 = @"^[0-9]{1,2}(\:[0-9]{1,2}){0,2}$";
    private const string RegEx12 = @"^[0-9]{1,2}(\:[0-9]{1,2}){0,2}( PM| AM)$";

    private static double GetSerialNumber(int hour, int minute, int second)
    {
        double secondsInADay = 24d * 60d * 60d;

        return (((double)hour * 60 * 60) + ((double)minute * 60) + (double)second) / secondsInADay;
    }

    private static void ValidateValues(int hour, int minute, int second)
    {
        if (second < 0 || second > 59)
        {
            throw new FormatException("Illegal value for second: " + second);
        }

        if (minute < 0 || minute > 59)
        {
            throw new FormatException("Illegal value for minute: " + minute);
        }
    }

    public virtual double Parse(string input) => InternalParse(input);

    public virtual bool CanParse(string input) => Regex.IsMatch(input, RegEx24) || Regex.IsMatch(input, RegEx12) || System.DateTime.TryParse(input, out System.DateTime _);

    private static double InternalParse(string input)
    {
        if (Regex.IsMatch(input, RegEx24))
        {
            return Parse24HourTimeString(input);
        }

        if (Regex.IsMatch(input, RegEx12))
        {
            return Parse12HourTimeString(input);
        }

        if (System.DateTime.TryParse(input, out System.DateTime dateTime))
        {
            return GetSerialNumber(dateTime.Hour, dateTime.Minute, dateTime.Second);
        }

        return -1;
    }

    private static double Parse12HourTimeString(string input)
    {
        string dayPart = input.Substring(input.Length - 2, 2);
        GetValuesFromString(input, out int hour, out int minute, out int second);

        if (dayPart == "PM")
        {
            hour += 12;
        }

        ValidateValues(hour, minute, second);

        return GetSerialNumber(hour, minute, second);
    }

    private static double Parse24HourTimeString(string input)
    {
        GetValuesFromString(input, out int hour, out int minute, out int second);
        ValidateValues(hour, minute, second);

        return GetSerialNumber(hour, minute, second);
    }

    private static void GetValuesFromString(string input, out int hour, out int minute, out int second)
    {
        minute = 0;
        second = 0;

        string[]? items = input.Split(':');
        hour = int.Parse(items[0]);

        if (items.Length > 1)
        {
            minute = int.Parse(items[1]);
        }

        if (items.Length > 2)
        {
            string? val = items[2];
            val = Regex.Replace(val, "[^0-9]+$", string.Empty);
            second = int.Parse(val);
        }
    }
}