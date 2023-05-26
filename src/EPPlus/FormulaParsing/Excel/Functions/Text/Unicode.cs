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
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "5.0",
        Description = "Returns the number (code point) corresponding to the first character of a supplied text string",
        IntroducedInExcelVersion = "2013")]
    internal class Unicode : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            string? arg = ArgToString(arguments, 0);
            if (!IsString(arg, allowNullOrEmpty: false))
            {
                return this.CreateResult(ExcelErrorValue.Values.Value, DataType.ExcelError);
            }

            string? firstChar = arg.Substring(0, 1);
            byte[]? bytes = Encoding.UTF32.GetBytes(firstChar);
            int code = BitConverter.ToInt32(bytes, 0);
            return CreateResult(code, DataType.Integer);
        }
    }
}
