using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

internal class CountIfsCompiler : FunctionCompiler
{
    public CountIfsCompiler(ExcelFunction function, ParsingContext context)
        : base(function, context)
    {
    }

    private readonly ExpressionEvaluator _evaluator = new ExpressionEvaluator();

    public override CompileResult Compile(IEnumerable<Expression> children)
    {
        List<FunctionArgument>? args = new List<FunctionArgument>();
        this.Function.BeforeInvoke(this.Context);

        for (int rangeIx = 0; rangeIx < children.Count(); rangeIx += 2)
        {
            Expression? rangeExpr = children.ElementAt(rangeIx).Children.First();
            rangeExpr.IgnoreCircularReference = true;
            RangeAddress? currentAdr = this.Context.Scopes.Current.Address;
            ExcelAddress? rangeAdr = new ExcelAddress(rangeExpr.ExpressionString);
            string? rangeWs = string.IsNullOrEmpty(rangeAdr.WorkSheetName) ? currentAdr.Worksheet : rangeAdr.WorkSheetName;

            if (currentAdr.Worksheet == rangeWs && rangeAdr.Collide(new ExcelAddress(currentAdr.Address)) != ExcelAddressBase.eAddressCollition.No)
            {
                object? candidateArg = children.ElementAt(rangeIx + 1)?.Children.FirstOrDefault()?.Compile().Result;

                if (children.ElementAt(rangeIx).HasChildren)
                {
                    int functionRowIndex = currentAdr.FromRow - rangeAdr._fromRow;
                    int functionColIndex = currentAdr.FromCol - rangeAdr._fromCol;
                    IRangeInfo? firstRangeResult = children.ElementAt(rangeIx).Children.First().Compile().Result as IRangeInfo;

                    if (firstRangeResult != null)
                    {
                        int candidateRowIndex = firstRangeResult.Address._fromRow + functionRowIndex;
                        int candidateColIndex = firstRangeResult.Address._fromCol + functionColIndex;
                        object? candidateValue = firstRangeResult.GetValue(candidateRowIndex, candidateColIndex);

                        if (this._evaluator.Evaluate(candidateArg, candidateValue?.ToString()))
                        {
                            if (this.Context.Configuration.AllowCircularReferences)
                            {
                                return CompileResult.ZeroDecimal;
                            }

                            throw new CircularReferenceException("Circular reference detected in " + currentAdr.Address);
                        }
                    }
                }
            }

            // todo: check circular ref for the actual cell where the SumIf formula resides (currentAdr).
        }

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

        return this.Function.Execute(args, this.Context);
    }
}