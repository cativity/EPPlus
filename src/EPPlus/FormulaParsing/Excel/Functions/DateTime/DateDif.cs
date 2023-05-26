/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/27/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "5.5",
        Description = "Get days, months, or years between two dates")]
    internal class DateDif : DateParsingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            object? startDateObj = arguments.ElementAt(0).Value;
            System.DateTime startDate = this.ParseDate(arguments, startDateObj);
            object? endDateObj = arguments.ElementAt(1).Value;
            System.DateTime endDate = this.ParseDate(arguments, endDateObj, 1);
            if (startDate > endDate)
            {
                return this.CreateResult(eErrorType.Num);
            }

            string? unit = ArgToString(arguments, 2);
            switch(unit.ToLower())
            {
                case "y":
                    return this.CreateResult(DateDiffYears(startDate, endDate), DataType.Integer);
                case "m":
                    return this.CreateResult(DateDiffMonths(startDate, endDate), DataType.Integer);
                case "d":
                    double daysD = endDate.Subtract(startDate).TotalDays;
                    return this.CreateResult(daysD, DataType.Integer);
                case "ym":
                    double monthsYm = DateDiffMonthsY(startDate, endDate);
                    return this.CreateResult(monthsYm, DataType.Integer);
                case "yd":
                    double daysYd = GetStartYearEndDateY(startDate, endDate).Subtract(startDate).TotalDays;
                    return this.CreateResult(daysYd, DataType.Integer);
                case "md":
                    // NB! Excel calculates wrong here sometimes. Example DATEDIF(2001-04-02, 2003-01-01, "md") = 30 (it should be 29)
                    // we have not implemented this bug in EPPlus. Microsoft advices not to use the DateDif function due to this and other bugs.
                    double daysMd = GetStartYearEndDateMd(startDate, endDate).Subtract(startDate).TotalDays;
                    return this.CreateResult(daysMd, DataType.Integer);
                default:
                    return this.CreateResult(eErrorType.Num);
            }
        }

        private static double DateDiffYears(System.DateTime start, System.DateTime end)
        {
            double result = Convert.ToDouble(end.Year - start.Year);
            System.DateTime tmpEnd = GetStartYearEndDate(start, end);
            if (start > tmpEnd)
            {
                result -= 1;
            }
            return result;
        }

        private static double DateDiffMonths(System.DateTime start, System.DateTime end)
        {
            double years = DateDiffYears(start, end);
            double result = years * 12;
            System.DateTime tmpEnd = GetStartYearEndDate(start, end);
            if(start > tmpEnd)
            {
                result += 12;
                while (start > tmpEnd)
                {
                    tmpEnd = tmpEnd.AddMonths(1);
                    result--;
                }
            }
            
            return result;
        }

        private static double DateDiffMonthsY(System.DateTime start, System.DateTime end)
        {
            System.DateTime endDate = GetStartYearEndDateY(start, end);
            double nMonths = 0d;
            System.DateTime tmpDate = start;
            if(tmpDate.AddMonths(1) < endDate)
            {
                do
                {
                    tmpDate = tmpDate.AddMonths(1);
                    if(tmpDate < endDate)
                    {
                        nMonths++;
                    }
                }
                while (tmpDate < endDate);
            }
            
            return nMonths;
        }

        private static System.DateTime GetStartYearEndDate(System.DateTime start, System.DateTime end)
        {
            return new System.DateTime(start.Year, end.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
        }

        private static System.DateTime GetStartYearEndDateY(System.DateTime start, System.DateTime end)
        {
            System.DateTime dt = new System.DateTime(start.Year, end.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
            if(dt < start)
            {
                dt = dt.AddYears(1);
            }
            return dt;
        }

        private static System.DateTime GetStartYearEndDateMd(System.DateTime start, System.DateTime end)
        {
            System.DateTime dt = new System.DateTime(start.Year, start.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
            if (dt < start)
            {
                dt = dt.AddMonths(1);
            }
            return dt;
        }
    }
}
