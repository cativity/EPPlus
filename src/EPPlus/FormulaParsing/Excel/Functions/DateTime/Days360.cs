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
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Implementations;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Calculates the number of days between 2 dates, based on a 360-day year (12 x 30 months)")]
    internal class Days360 : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            double numDate1 = ArgToDecimal(arguments, 0);
            double numDate2 = ArgToDecimal(arguments, 1);
            System.DateTime dt1 = System.DateTime.FromOADate(numDate1);
            System.DateTime dt2 = System.DateTime.FromOADate(numDate2);

            Days360Calctype calcType = Days360Calctype.Us;
            if (arguments.Count() > 2)
            {
                bool european = ArgToBool(arguments, 2);
                if (european)
                {
                    calcType = Days360Calctype.European;
                }
            }

            int result = Days360Impl.CalcDays360(dt1, dt2, calcType);
            return CreateResult(result, DataType.Integer);
        }

        private int GetNumWholeMonths(System.DateTime dt1, System.DateTime dt2)
        {
            System.DateTime startDate = new System.DateTime(dt1.Year, dt1.Month, 1).AddMonths(1);
            System.DateTime endDate = new System.DateTime(dt2.Year, dt2.Month, 1);
            return ((endDate.Year - startDate.Year)*12) + (endDate.Month - startDate.Month);
        }
    }
}
