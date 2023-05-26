using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    internal class CountIfFunctionCompiler : FunctionCompiler
    {
        public CountIfFunctionCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
        }

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            List<FunctionArgument>? args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            foreach (Expression? child in children)
            {
                CompileResult? compileResult = child.Compile();
                if (compileResult.IsResultOfSubtotal)
                {
                    FunctionArgument? arg = new FunctionArgument(compileResult.Result, compileResult.DataType);
                    arg.SetExcelStateFlag(ExcelCellState.IsResultOfSubtotal);
                    args.Add(arg);
                }
                else
                {
                    BuildFunctionArguments(compileResult, args);
                }
            }
            if (args.Count < 2)
            {
                return new CompileResult(eErrorType.Value);
            }

            FunctionArgument? arg2 = args.ElementAt(1);
            if(arg2.DataType == DataType.Enumerable && arg2.IsExcelRange)
            {
                FunctionArgument? arg1 = args.First();
                List<object>? result = new List<object>();
                IRangeInfo? rangeValues = arg2.ValueAsRangeInfo;
                foreach(ICellInfo? funcArg in rangeValues)
                {
                    List<FunctionArgument>? arguments = new List<FunctionArgument> { arg1 };
                    CompileResult? cr = new CompileResultFactory().Create(funcArg.Value);
                    BuildFunctionArguments(cr, arguments);
                    CompileResult? r = Function.Execute(arguments, Context);
                    result.Add(r.Result);
                }
                return new CompileResult(result, DataType.Enumerable);
            }
            return Function.Execute(args, Context);
        }
    }
}
