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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

[FunctionMetadata(Category = ExcelFunctionCategory.DateAndTime,
                  EPPlusVersion = "4",
                  Description =
                      "Returns a date that is the last day of the month which is a specified number of months before or after an initial supplied start date")]
internal class Eomonth : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        System.DateTime date = System.DateTime.FromOADate(this.ArgToDecimal(arguments, 0));
        int monthsToAdd = this.ArgToInt(arguments, 1);
        System.DateTime resultDate = new System.DateTime(date.Year, date.Month, 1).AddMonths(monthsToAdd + 1).AddDays(-1);

        return this.CreateResult(resultDate.ToOADate(), DataType.Date);
    }
}