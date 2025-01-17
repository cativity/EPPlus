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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.MathAndTrig,
                  EPPlusVersion = "4",
                  Description = "Returns the sum of the squares of a supplied list of numbers")]
internal class Sumsq : HiddenValuesHandlingFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        double retVal = 0d;

        if (arguments != null)
        {
            foreach (FunctionArgument? arg in arguments)
            {
                retVal += this.Calculate(arg, context);
            }
        }

        return this.CreateResult(retVal, DataType.Decimal);
    }

    private double Calculate(FunctionArgument arg, ParsingContext context, bool isInArray = false)
    {
        double retVal = 0d;

        if (this.ShouldIgnore(arg, context))
        {
            return retVal;
        }

        if (arg.Value is IEnumerable<FunctionArgument>)
        {
            foreach (FunctionArgument? item in (IEnumerable<FunctionArgument>)arg.Value)
            {
                retVal += this.Calculate(item, context, true);
            }
        }
        else
        {
            IRangeInfo? cs = arg.Value as IRangeInfo;

            if (cs != null)
            {
                foreach (ICellInfo? c in cs)
                {
                    if (this.ShouldIgnore(c, context) == false)
                    {
                        CheckForAndHandleExcelError(c);
                        retVal += System.Math.Pow(c.ValueDouble, 2);
                    }
                }
            }
            else
            {
                CheckForAndHandleExcelError(arg);

                if (IsNumericString(arg.Value) && !isInArray)
                {
                    double value = ConvertUtil.GetValueDouble(arg.Value);

                    return System.Math.Pow(value, 2);
                }

                retVal += System.Math.Pow(ConvertUtil.GetValueDouble(arg.Value, isInArray), 2);
            }
        }

        return retVal;
    }
}