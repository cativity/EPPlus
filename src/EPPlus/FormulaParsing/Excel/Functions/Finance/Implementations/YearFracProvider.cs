using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class YearFracProvider : IYearFracProvider
    {
        public YearFracProvider(ParsingContext context)
        {
            this._context = context;
        }

        private readonly ParsingContext _context;
        public double GetYearFrac(System.DateTime date1, System.DateTime date2, DayCountBasis basis)
        {
            Yearfrac? func = new Yearfrac();
            List<FunctionArgument>? args = new List<FunctionArgument> { new FunctionArgument(date1.ToOADate()), new FunctionArgument(date2.ToOADate()), new FunctionArgument((int)basis) };
            CompileResult? result = func.Execute(args, this._context);
            return result.ResultNumeric;
        }
    }
}
