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
        Description = "Calculates the discount rate for a security")]
    internal class Disc : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            double settlementNum = ArgToDecimal(arguments, 0);
            double maturityNum = ArgToDecimal(arguments, 1);
            System.DateTime settlement = System.DateTime.FromOADate(settlementNum);
            System.DateTime maturity = System.DateTime.FromOADate(maturityNum);
            double pr = ArgToDecimal(arguments, 2);
            double redemption = ArgToDecimal(arguments, 3);
            int basis = 0;
            if(arguments.Count() > 4)
            {
                basis = ArgToInt(arguments, 4);
            }
            if(maturity <= settlement || pr <= 0 || redemption <= 0 || (basis < 0 || basis > 4))
            {
                return CreateResult(eErrorType.Num);
            }
            YearFracProvider? yearFrac = new YearFracProvider(context);
            double result = (1d - pr / redemption) / yearFrac.GetYearFrac(settlement, maturity, (DayCountBasis)basis);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
