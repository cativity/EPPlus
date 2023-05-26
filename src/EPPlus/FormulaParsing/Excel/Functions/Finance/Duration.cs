/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.2",
        Description = "Calculates the Macauley duration of a security with an assumed par value of $100")]
    internal class Duration : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 5);
            double settlementNum = this.ArgToDecimal(arguments, 0);
            double maturityNum = this.ArgToDecimal(arguments, 1);
            System.DateTime settlement = System.DateTime.FromOADate(settlementNum);
            System.DateTime maturity = System.DateTime.FromOADate(maturityNum);
            double coupon = this.ArgToDecimal(arguments, 2);
            double yield = this.ArgToDecimal(arguments, 3);
            if(coupon < 0 || yield < 0)
            {
                return this.CreateResult(eErrorType.Num);
            }
            int frequency = this.ArgToInt(arguments, 4);
            if(frequency != 1 && frequency != 2 && frequency != 4)
            {
                return this.CreateResult(eErrorType.Num);
            }
            int basis = 0;
            if(arguments.Count() > 5)
            {
                basis = this.ArgToInt(arguments, 5);
            }
            if(basis < 0 || basis > 4)
            {
                return this.CreateResult(eErrorType.Num);
            }
            DurationImpl? func = new DurationImpl(new YearFracProvider(context), new CouponProvider());
            double result = func.GetDuration(settlement, maturity, coupon, yield, frequency, (DayCountBasis)basis);
            return this.CreateResult(result, DataType.Decimal);
        }
    }
}
