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
using System.Security.Cryptography;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays
{
    internal class WorkdayCalculator
    {
        private readonly HolidayWeekdays _holidayWeekdays;

        public WorkdayCalculator()
            : this(new HolidayWeekdays())
        {}

        public WorkdayCalculator(HolidayWeekdays holidayWeekdays)
        {
            this._holidayWeekdays = holidayWeekdays;
        }

        public WorkdayCalculatorResult CalculateNumberOfWorkdays(System.DateTime startDate, System.DateTime endDate)
        {
            WorkdayCalculationDirection calcDirection = startDate < endDate
                                                            ? WorkdayCalculationDirection.Forward
                                                            : WorkdayCalculationDirection.Backward;
            System.DateTime calcStartDate;
            System.DateTime calcEndDate;
            if (calcDirection == WorkdayCalculationDirection.Forward)
            {
                calcStartDate = startDate.Date;
                calcEndDate = endDate.Date;
            }
            else
            {
                calcStartDate = endDate.Date;
                calcEndDate = startDate.Date;
            }
            int nWholeWeeks = (int)calcEndDate.Subtract(calcStartDate).TotalDays/7;
            int workdaysCounted = nWholeWeeks* this._holidayWeekdays.NumberOfWorkdaysPerWeek;
            if (!this._holidayWeekdays.IsHolidayWeekday(calcStartDate))
            {
                workdaysCounted++;
            }
            System.DateTime tmpDate = calcStartDate.AddDays(nWholeWeeks*7);
            while (tmpDate < calcEndDate)
            {
                tmpDate = tmpDate.AddDays(1);
                if (!this._holidayWeekdays.IsHolidayWeekday(tmpDate))
                {
                    workdaysCounted++;
                }
            }
            return new WorkdayCalculatorResult(workdaysCounted, startDate, endDate, calcDirection);
        }

        public WorkdayCalculatorResult CalculateWorkday(System.DateTime startDate, int nWorkDays)
        {
            WorkdayCalculationDirection calcDirection = nWorkDays > 0 ? WorkdayCalculationDirection.Forward : WorkdayCalculationDirection.Backward;
            int direction = (int) calcDirection;
            nWorkDays *= direction;
            int workdaysCounted = 0;
            System.DateTime tmpDate = startDate;
            
            // calculate whole weeks
            int nWholeWeeks = nWorkDays / this._holidayWeekdays.NumberOfWorkdaysPerWeek;
            tmpDate = tmpDate.AddDays(nWholeWeeks * 7 * direction);
            workdaysCounted += nWholeWeeks * this._holidayWeekdays.NumberOfWorkdaysPerWeek;

            // calculate the rest
            while (workdaysCounted < nWorkDays)
            {
                tmpDate = tmpDate.AddDays(direction);
                if (!this._holidayWeekdays.IsHolidayWeekday(tmpDate))
                {
                    workdaysCounted++;
                }
            }
            return new WorkdayCalculatorResult(workdaysCounted, startDate, tmpDate, calcDirection);
        }

        public WorkdayCalculatorResult ReduceWorkdaysWithHolidays(WorkdayCalculatorResult calculatedResult,
            FunctionArgument holidayArgument)
        {
            System.DateTime startDate = calculatedResult.StartDate;
            System.DateTime endDate = calculatedResult.EndDate;
            AdditionalHolidayDays? additionalDays = new AdditionalHolidayDays(holidayArgument);
            System.DateTime calcStartDate;
            System.DateTime calcEndDate;
            if (startDate < endDate)
            {
                calcStartDate = startDate;
                calcEndDate = endDate;
            }
            else
            {
                calcStartDate = endDate;
                calcEndDate = startDate;
            }
            int nAdditionalHolidayDays = additionalDays.AdditionalDates.Count(x => x >= calcStartDate && x <= calcEndDate && !this._holidayWeekdays.IsHolidayWeekday(x));
            return new WorkdayCalculatorResult(calculatedResult.NumberOfWorkdays - nAdditionalHolidayDays, startDate, endDate, calculatedResult.Direction);
        } 

        public WorkdayCalculatorResult AdjustResultWithHolidays(WorkdayCalculatorResult calculatedResult,
                                                         FunctionArgument holidayArgument)
        {
            System.DateTime startDate = calculatedResult.StartDate;
            System.DateTime endDate = calculatedResult.EndDate;
            WorkdayCalculationDirection direction = calculatedResult.Direction;
            int workdaysCounted = calculatedResult.NumberOfWorkdays;
            AdditionalHolidayDays? additionalDays = new AdditionalHolidayDays(holidayArgument);
            foreach (System.DateTime date in additionalDays.AdditionalDates)
            {
                if (direction == WorkdayCalculationDirection.Forward && (date < startDate || date > endDate))
                {
                    continue;
                }

                if (direction == WorkdayCalculationDirection.Backward && (date > startDate || date < endDate))
                {
                    continue;
                }

                if (this._holidayWeekdays.IsHolidayWeekday(date))
                {
                    continue;
                }

                System.DateTime tmpDate = this._holidayWeekdays.GetNextWorkday(endDate, direction);
                while (additionalDays.AdditionalDates.Contains(tmpDate))
                {
                    tmpDate = this._holidayWeekdays.GetNextWorkday(tmpDate, direction);
                }
                workdaysCounted++;
                endDate = tmpDate;
            }

            return new WorkdayCalculatorResult(workdaysCounted, calculatedResult.StartDate, endDate, direction);
        }
    }
}
