/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Implementations;

internal static class Days360Impl
{
    internal static int CalcDays360(System.DateTime startDate, System.DateTime endDate, Days360Calctype calcType)
    {
        int startYear = startDate.Year;
        int startMonth = startDate.Month;
        int startDay = startDate.Day;
        int endYear = endDate.Year;
        int endMonth = endDate.Month;
        int endDay = endDate.Day;

        if (calcType == Days360Calctype.European)
        {
            if (startDay == 31)
            {
                startDay = 30;
            }

            if (endDay == 31)
            {
                endDay = 30;
            }
        }
        else
        {
            GregorianCalendar? calendar = new GregorianCalendar();
            int nDaysInFeb = calendar.IsLeapYear(startDate.Year) ? 29 : 28;

            // If the investment is EOM and (Date1 is the last day of February) and (Date2 is the last day of February), then change D2 to 30.
            if (startMonth == 2 && startDay == nDaysInFeb && endMonth == 2 && endDay == nDaysInFeb)
            {
                endDay = 30;
            }

            // If the investment is EOM and (Date1 is the last day of February), then change D1 to 30.
            if (startMonth == 2 && startDay == nDaysInFeb)
            {
                startDay = 30;
            }

            // If D2 is 31 and D1 is 30 or 31, then change D2 to 30.
            if (endDay == 31 && (startDay == 30 || startDay == 31))
            {
                endDay = 30;
            }

            // If D1 is 31, then change D1 to 30.
            if (startDay == 31)
            {
                startDay = 30;
            }
        }

        int result = (endYear * 12 * 30) + (endMonth * 30) + endDay - ((startYear * 12 * 30) + (startMonth * 30) + startDay);

        return result;
    }
}