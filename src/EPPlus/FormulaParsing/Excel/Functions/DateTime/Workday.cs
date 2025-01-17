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
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

[FunctionMetadata(Category = ExcelFunctionCategory.DateAndTime,
                  EPPlusVersion = "4",
                  Description = "Returns a date that is a supplied number of working days (excluding weekends & holidays) ahead of a given start date")]
internal class Workday : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        FunctionArgument[]? functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
        ValidateArguments(functionArguments, 2);
        System.DateTime startDate = System.DateTime.FromOADate(this.ArgToInt(functionArguments, 0));
        int nWorkDays = this.ArgToInt(functionArguments, 1);

        WorkdayCalculator? calculator = new WorkdayCalculator();
        WorkdayCalculatorResult? result = calculator.CalculateWorkday(startDate, nWorkDays);

        if (functionArguments.Length > 2)
        {
            result = calculator.AdjustResultWithHolidays(result, functionArguments[2]);
        }

        return this.CreateResult(result.EndDate.ToOADate(), DataType.Date);
    }
}