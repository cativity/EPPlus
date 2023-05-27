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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.MathAndTrig,
                  EPPlusVersion = "5.0",
                  Description = "Rounds a number away from zero (i.e. rounds a positive number up and a negative number down), to the next even number")]
internal class Even : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        object? arg = arguments.ElementAt(0).Value;

        if (!IsNumeric(arg))
        {
            return new CompileResult(eErrorType.Value);
        }

        double number = ConvertUtil.GetValueDouble(arg);

        if (number % 1 != 0)
        {
            if (number >= 0)
            {
                number = number - (number % 1) + 1;
            }
            else
            {
                number = number - (number % 1) - 1;
            }
        }

        int intNumber = Convert.ToInt32(number);

        if (intNumber % 2 == 1)
        {
            intNumber = intNumber >= 0 ? intNumber + 1 : intNumber - 1;
        }

        return this.CreateResult(intNumber, DataType.Integer);
    }
}