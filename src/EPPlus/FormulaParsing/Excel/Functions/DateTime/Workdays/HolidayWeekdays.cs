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
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays
{
    public class HolidayWeekdays
    {
        private readonly List<DayOfWeek> _holidayDays = new List<DayOfWeek>();

        public HolidayWeekdays()
            :this(DayOfWeek.Saturday, DayOfWeek.Sunday)
        {
            
        }

        public int NumberOfWorkdaysPerWeek => 7 - this._holidayDays.Count;

        public HolidayWeekdays(params DayOfWeek[] holidayDays)
        {
            foreach (DayOfWeek dayOfWeek in holidayDays)
            {
                this._holidayDays.Add(dayOfWeek);
            }
        }

        public bool IsHolidayWeekday(System.DateTime dateTime)
        {
            return this._holidayDays.Contains(dateTime.DayOfWeek);
        }

        public System.DateTime AdjustResultWithHolidays(System.DateTime resultDate,
                                                         IEnumerable<FunctionArgument> arguments)
        {
            if (arguments.Count() == 2)
            {
                return resultDate;
            }

            IEnumerable<FunctionArgument>? holidays = arguments.ElementAt(2).Value as IEnumerable<FunctionArgument>;
            if (holidays != null)
            {
                foreach (FunctionArgument? arg in holidays)
                {
                    if (ConvertUtil.IsNumericOrDate(arg.Value))
                    {
                        double dateSerial = ConvertUtil.GetValueDouble(arg.Value);
                        System.DateTime holidayDate = System.DateTime.FromOADate(dateSerial);
                        if (!this.IsHolidayWeekday(holidayDate))
                        {
                            resultDate = resultDate.AddDays(1);
                        }
                    }
                }
            }
            else
            {
                IRangeInfo? range = arguments.ElementAt(2).Value as IRangeInfo;
                if (range != null)
                {
                    foreach (ICellInfo? cell in range)
                    {
                        if (ConvertUtil.IsNumericOrDate(cell.Value))
                        {
                            double dateSerial = ConvertUtil.GetValueDouble(cell.Value);
                            System.DateTime holidayDate = System.DateTime.FromOADate(dateSerial);
                            if (!this.IsHolidayWeekday(holidayDate))
                            {
                                resultDate = resultDate.AddDays(1);
                            }
                        }
                    }
                }
            }
            return resultDate;
        }

        public System.DateTime GetNextWorkday(System.DateTime date, WorkdayCalculationDirection direction = WorkdayCalculationDirection.Forward)
        {
            int changeParam = (int)direction;
            System.DateTime tmpDate = date.AddDays(changeParam);
            while (this.IsHolidayWeekday(tmpDate))
            {
                tmpDate = tmpDate.AddDays(changeParam);
            }
            return tmpDate;
        }
    }
}
