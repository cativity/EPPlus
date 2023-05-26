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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Statistical,
                     EPPlusVersion = "4",
                     Description = "Returns the statistical rank of a given value, within a supplied array of values")]
internal class Rank : RankFunctionBase
{
    public Rank()
        : this(false)
    {

    }
        
    public Rank(bool isAvg)
    {
        this._isAvg=isAvg;
    }

    private readonly bool _isAvg;

    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        double number = this.ArgToDecimal(arguments, 0);
        FunctionArgument? refArg = arguments.ElementAt(1);
        bool sortAscending = arguments.Count() > 2 ? this.ArgToBool(arguments, 2) : false;
        List<double>? numbers = GetNumbersFromRange(refArg, sortAscending);
        double rank = numbers.IndexOf(number) + 1;
        if(this._isAvg)
        {
            int lastRank = numbers.LastIndexOf(number) + 1;
            rank += ((lastRank - rank) / 2d);
        }
            
        if (rank <= 0 || rank > numbers.Count)
        {
            return new CompileResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
        }
        return this.CreateResult(rank, DataType.Decimal);
    }

        
}