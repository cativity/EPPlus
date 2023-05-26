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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Calculates the number of days between 2 dates",
        IntroducedInExcelVersion = "2013")]
    public class Days : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            ValidateArguments(arguments, 2);
            double numDate1 = this.ArgToDecimal(arguments, 0);
            double numDate2 = this.ArgToDecimal(arguments, 1);
            System.DateTime endDate = System.DateTime.FromOADate(numDate1);
            System.DateTime startDate = System.DateTime.FromOADate(numDate2);
            return this.CreateResult(endDate.Subtract(startDate).TotalDays, DataType.Date);
        }
    }
}
