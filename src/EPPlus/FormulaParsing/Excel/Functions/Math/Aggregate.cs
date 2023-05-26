/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
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
         IntroducedInExcelVersion = "2010",
         Description = "Performs a specified calculation (e.g. the sum, product, average, etc.) for a list or database, with the option to ignore hidden rows and error values")]
    internal class Aggregate : ExcelFunction
    {
        public override void BeforeInvoke(ParsingContext context)
        {
            base.BeforeInvoke(context);
            ulong cellId = context.ExcelDataProvider.GetCellId(context.Scopes.Current.Address.Worksheet, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol);
            if (!context.SubtotalAddresses.Contains(cellId))
            {
                context.SubtotalAddresses.Add(cellId);
            }
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            int funcNum = ArgToInt(arguments, 0);
            int nToSkip = IsNumeric(arguments.ElementAt(1).Value) ? 2 : 1;  
            int options = nToSkip == 1 ? 0 : ArgToInt(arguments, 1);

            if (options < 0 || options > 7)
            {
                return this.CreateResult(eErrorType.Value);
            }

            if(IgnoreNestedSubtotalAndAggregate(options))
            {
                context.Scopes.Current.IsSubtotal = true;
                ulong cellId = context.ExcelDataProvider.GetCellId(context.Scopes.Current.Address.Worksheet, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol);
                if(!context.SubtotalAddresses.Contains(cellId))
                {
                    context.SubtotalAddresses.Add(cellId);
                }
            }

            CompileResult result = null;
            switch(funcNum)
            {
                case 1:
                    Average? f1 = new Average()
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f1.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 2:
                    Count? f2 = new Count()
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f2.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 3:
                    CountA? f3 = new CountA
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f3.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 4:
                    Max? f4 = new Max 
                    { 
                        IgnoreHiddenValues = IgnoreHidden(options), 
                        IgnoreErrors = IgnoreErrors(options) 
                    };
                    result = f4.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 5:
                    Min? f5 = new Min
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f5.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 6:
                    Product? f6 = new Product
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f6.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 7:
                    StdevDotS? f7 = new StdevDotS
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f7.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 8:
                    StdevDotP? f8 = new StdevDotP
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f8.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 9:
                    Sum? f9 = new Sum
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f9.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 10:
                    VarDotS f10 = new VarDotS
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f10.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 11:
                    VarDotP? f11 = new VarDotP
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f11.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 12:
                    Median? f12 = new Median
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f12.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 13:
                    ModeSngl? f13 = new ModeSngl
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f13.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 14:
                    Large? f14 = new Large
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    FunctionArgument? a141 = arguments.ElementAt(nToSkip);
                    FunctionArgument? a142 = arguments.ElementAt(nToSkip + 1);
                    result = f14.Execute(new List<FunctionArgument> { a141, a142 }, context);
                    break;
                case 15:
                    Small? f15 = new Small
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f15.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                case 16:
                    PercentileInc? f16 = new PercentileInc
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f16.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                case 17:
                    QuartileInc? f17 = new QuartileInc
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f17.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                case 18:
                    PercentileExc? f18 = new PercentileExc
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f18.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                case 19:
                    QuartileExc? f19 = new QuartileExc
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f19.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                default:
                    result = CreateResult(eErrorType.Value);
                    break;
            }
            result.IsResultOfSubtotal = IgnoreNestedSubtotalAndAggregate(options);
            return result;
        }

        private static bool IgnoreHidden(int options)
        {
            return options == 1 || options == 3 || options == 5 || options == 7;
        }

        private static bool IgnoreErrors(int options)
        {
            return options == 2 || options == 3 || options == 6 || options == 7;
        }

        private static bool IgnoreNestedSubtotalAndAggregate(int options)
        {
            return options == 0 || options == 1 || options == 2 || options == 3;
        }
    }
}
