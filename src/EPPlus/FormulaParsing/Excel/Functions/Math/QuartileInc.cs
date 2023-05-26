using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.Statistical,
            EPPlusVersion = "5.5",
            Description = "Returns the K'th percentile of values in a supplied range, where K is in the range 0 - 1 (inclusive)")]
    internal class QuartileInc : PercentileInc
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            IEnumerable<FunctionArgument>? arrArg = arguments.Take(1);
            List<double>? arr = this.ArgsToDoubleEnumerable(arrArg, context).Select(x => (double)x).ToList();
            if (!arr.Any())
            {
                return this.CreateResult(eErrorType.Value);
            }

            int quart = this.ArgToInt(arguments, 1);
            switch (quart)
            {
                case 0:
                    return this.CreateResult(arr.Min(), DataType.Decimal);
                case 1:
                    return base.Execute(BuildArgs(arrArg, 0.25d), context);
                case 2:
                    return base.Execute(BuildArgs(arrArg, 0.5d), context);
                case 3:
                    return base.Execute(BuildArgs(arrArg, 0.75d), context);
                case 4:
                    return this.CreateResult(arr.Max(), DataType.Decimal);
                default:
                    return this.CreateResult(eErrorType.Num);
            }
        }

        private static IEnumerable<FunctionArgument> BuildArgs(IEnumerable<FunctionArgument> arrArg, double quart)
        {
            List<FunctionArgument>? argList = new List<FunctionArgument>();
            argList.AddRange(arrArg);
            argList.Add(new FunctionArgument(quart, DataType.Decimal));
            return argList;
        }
    }
}
