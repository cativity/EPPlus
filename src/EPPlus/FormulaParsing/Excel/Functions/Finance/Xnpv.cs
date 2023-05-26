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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Financial,
       EPPlusVersion = "5.2",
       Description = "Calculates the net present value for a schedule of cash flows occurring at a series of supplied dates")]
    internal class Xnpv : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            double rate = ArgToDecimal(arguments, 0);
            List<FunctionArgument>? arg2 = new List<FunctionArgument> { arguments.ElementAt(1) };
            IEnumerable<ExcelDoubleCellValue>? values = ArgsToDoubleEnumerable(arg2, context);
            IEnumerable<System.DateTime>? dates = GetDates(arguments.ElementAt(2), context);
            if (values.Count() != dates.Count())
            {
                return this.CreateResult(eErrorType.Num);
            }

            System.DateTime firstDate = dates.First();
            double result = 0d;
            for(int i = 0; i < values.Count(); i++)
            {
                System.DateTime dt = dates.ElementAt(i);
                ExcelDoubleCellValue val = values.ElementAt(i);
                if (dt < firstDate)
                {
                    return this.CreateResult(eErrorType.Num);
                }

                result += val / System.Math.Pow(1d + rate, dt.Subtract(firstDate).TotalDays / 365d);
            }
            return CreateResult(result, DataType.Decimal);
        }

        private static IEnumerable<System.DateTime> GetDates(FunctionArgument arg, ParsingContext context)
        {
            List<System.DateTime>? dates = new List<System.DateTime>();
            if(arg.Value is IEnumerable<FunctionArgument>)
            {
                IEnumerable<int>? args = ((IEnumerable<FunctionArgument>)arg.Value).Select(x => (int)x.Value);
                foreach(int num in args)
                {
                    dates.Add(System.DateTime.FromOADate(num));
                }
            }
            else if (arg.Value is IRangeInfo)
            {
                foreach (ICellInfo? c in (IRangeInfo)arg.Value)
                {
                    int num = Convert.ToInt32(c.ValueDouble);
                    dates.Add(System.DateTime.FromOADate(num));
                }
            }
            return dates;
        }
    }
}
