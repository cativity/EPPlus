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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Returns a time, from a user-supplied hour, minute and second")]
    internal class Time : TimeBaseFunction
    {
        public Time()
            : base()
        {

        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            string? firstArg = arguments.ElementAt(0).Value.ToString();
            if(arguments.Count() == 1 && this.TimeStringParser.CanParse(firstArg))
            {
                double result = this.TimeStringParser.Parse(firstArg);
                return new CompileResult(result, DataType.Time);
            }
            ValidateArguments(arguments, 3);
            int hour = this.ArgToInt(arguments, 0);
            int min = this.ArgToInt(arguments, 1);
            int sec = this.ArgToInt(arguments, 2);

            if (sec < 0 || sec > 59)
            {
                return this.CreateResult(eErrorType.Value);
            }

            if (min < 0 || min > 59)
            {
                return this.CreateResult(eErrorType.Value);
            }

            if (min < 0 || hour > 23)
            {
                return this.CreateResult(eErrorType.Value);
            }

            double secondsOfThisTime = (double)(hour * 60 * 60 + min * 60 + sec);
            return this.CreateResult(GetTimeSerialNumber(secondsOfThisTime), DataType.Time);
        }
    }
}
