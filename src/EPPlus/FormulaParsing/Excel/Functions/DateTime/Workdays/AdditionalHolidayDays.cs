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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;

public class AdditionalHolidayDays
{
    private readonly FunctionArgument _holidayArg;
    private readonly List<System.DateTime> _holidayDates = new List<System.DateTime>();

    public AdditionalHolidayDays(FunctionArgument holidayArg)
    {
        this._holidayArg = holidayArg;
        this.Initialize();
    }

    public IEnumerable<System.DateTime> AdditionalDates => this._holidayDates;

    private void Initialize()
    {
        IEnumerable<FunctionArgument>? holidays = this._holidayArg.Value as IEnumerable<FunctionArgument>;

        if (holidays != null)
        {
            foreach (System.DateTime holidayDate in from arg in holidays
                                                    where ConvertUtil.IsNumericOrDate(arg.Value)
                                                    select ConvertUtil.GetValueDouble(arg.Value)
                                                    into dateSerial
                                                    select System.DateTime.FromOADate(dateSerial))
            {
                this._holidayDates.Add(holidayDate);
            }
        }

        IRangeInfo? range = this._holidayArg.Value as IRangeInfo;

        if (range != null)
        {
            foreach (System.DateTime holidayDate in from cell in range
                                                    where ConvertUtil.IsNumericOrDate(cell.Value)
                                                    select ConvertUtil.GetValueDouble(cell.Value)
                                                    into dateSerial
                                                    select System.DateTime.FromOADate(dateSerial))
            {
                this._holidayDates.Add(holidayDate);
            }
        }

        if (ConvertUtil.IsNumericOrDate(this._holidayArg.Value))
        {
            this._holidayDates.Add(System.DateTime.FromOADate(ConvertUtil.GetValueDouble(this._holidayArg.Value)));
        }
    }
}