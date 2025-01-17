﻿/*************************************************************************************************
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
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical,
                  EPPlusVersion = "5.5",
                  IntroducedInExcelVersion = "2010",
                  Description = "Returns the specified quartile of a set of supplied numbers, based on percentile value 0 - 1 (exclusive) ")]
internal class QuartileExc : PercentileExc
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
            case 1:
                return base.Execute(BuildArgs(arrArg, 0.25d), context);

            case 2:
                return base.Execute(BuildArgs(arrArg, 0.5d), context);

            case 3:
                return base.Execute(BuildArgs(arrArg, 0.75d), context);

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