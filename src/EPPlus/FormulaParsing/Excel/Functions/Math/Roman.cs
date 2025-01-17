﻿/*************************************************************************************************
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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Math.RomanFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.MathAndTrig,
                  EPPlusVersion = "5.1",
                  Description = "Returns a text string depicting the roman numeral for a given number")]
internal class Roman : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        int number = this.ArgToInt(arguments, 0, RoundingMethod.Floor);
        int type = arguments.Count() > 1 ? this.FirstArgumentToInt(arguments) : 0;

        if (type < 0 || type > 4)
        {
            return this.CreateResult(eErrorType.Value);
        }

        if (number < 0 || number > 3999)
        {
            return this.CreateResult(eErrorType.Value);
        }

        RomanBase func = new RomanClassic();

        switch (type)
        {
            case 1:
                func = new RomanForm1();

                break;

            case 2:
                func = new RomanForm2();

                break;

            case 3:
                func = new RomanForm3();

                break;

            case 4:
                func = new RomanSimplified();

                break;

            default:
                break;
        }

        return this.CreateResult(func.Execute(number), DataType.String);
    }

    private int FirstArgumentToInt(IEnumerable<FunctionArgument> arguments)
    {
        FunctionArgument? arg = arguments.ElementAt(1);

        if (arg.DataType == DataType.Boolean && arg.ValueFirst is bool boolValue)
        {
            return boolValue ? 0 : 4;
        }

        return this.ArgToInt(arguments, 1);
    }
}