/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/03/2021         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.5",
        Description = "Calculates the depreciation of an asset for a specified period, using the fixed-declining balance method")]
    internal class Db : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            double cost = this.ArgToDecimal(arguments, 0);
            double salvage = this.ArgToDecimal(arguments, 1);
            double life = this.ArgToDecimal(arguments, 2);
            double period = this.ArgToDecimal(arguments, 3);
            int month = 12;
            if (arguments.Count() >= 5)
            {
                month = this.ArgToInt(arguments, 4);
            }

            if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month <= 0 || month > 12)
            {
                return this.CreateResult(eErrorType.Num);
            }

            if (period > life && month == 12 || period > (life + 1))
            {
                return this.CreateResult(eErrorType.Num);
            }

            // calculations below as described at https://support.microsoft.com/en-us/office/db-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7?ui=en-us&rs=en-us&ad=us

            // rate should be rounded to three decimals
            double rate = (1 - System.Math.Pow(salvage / cost, 1 / life));
            rate = System.Math.Round(rate, 3);

            // calculate first period
            double firstDepr = cost * rate * month / 12;

            if (period == 1)
            {
                return this.CreateResult(firstDepr, DataType.Decimal);
            }

            // remaining periods
            double total = firstDepr;
            double currentPeriodDepr = 0d;
            double toPeriod = (period == life) ? life - 1 : period;
            for (int i = 2; i <= toPeriod; i++)
            {
                currentPeriodDepr = (cost - total) * rate;
                total += currentPeriodDepr;
            }

            // Special case for the last period
            if (period >= life)
            {
                double result = (cost - total) * rate;
                if(period > life)
                {
                    // For the last period, DB uses this formula: ((cost - total depreciation from prior periods) * rate * (12 - month)) / 12
                    result = currentPeriodDepr  * (12 - month) / 12;
                }
                return this.CreateResult(result, DataType.Decimal);
            }
            else
            {
                return this.CreateResult(currentPeriodDepr, DataType.Decimal);
            }
        }
    }
}
