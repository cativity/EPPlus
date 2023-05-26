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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Replaces all or part of a text string with another string (from a user supplied position)")]
    internal class Replace : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            string? oldText = ArgToString(arguments, 0);
            int startPos = ArgToInt(arguments, 1);
            int nCharsToReplace = ArgToInt(arguments, 2);
            string? newText = ArgToString(arguments, 3);
            string? firstPart = GetFirstPart(oldText, startPos);
            string? lastPart = GetLastPart(oldText, startPos, nCharsToReplace);
            string? result = string.Concat(firstPart, newText, lastPart);
            return CreateResult(result, DataType.String);
        }

        private static string GetFirstPart(string text, int startPos)
        {
            return text.Substring(0, startPos - 1);
        }

        private static string GetLastPart(string text, int startPos, int nCharactersToReplace)
        {
            int startIx = startPos -1;
            startIx += nCharactersToReplace;
            return text.Substring(startIx, text.Length - startIx);
        }
    }
}
