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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of numerical values in a supplied set of cells or values")]
    internal class Count : HiddenValuesHandlingFunction
    {
        private enum ItemContext
        {
            InRange,
            InArray,
            SingleArg
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double nItems = 0d;
            this.Calculate(arguments, ref nItems, context, ItemContext.SingleArg);
            return this.CreateResult(nItems, DataType.Integer);
        }

        private void Calculate(IEnumerable<FunctionArgument> items, ref double nItems, ParsingContext context, ItemContext itemContext)
        {
            foreach (FunctionArgument? item in items)
            {
                IRangeInfo? cs = item.Value as IRangeInfo;
                if (cs != null)
                {
                    foreach (ICellInfo? c in cs)
                    {
                        _CheckForAndHandleExcelError(c, context);
                        if (this.ShouldIgnore(c, context) == false && ShouldCount(c.Value, ItemContext.InRange))
                        {
                            nItems++;
                        }
                    }
                }
                else
                {
                    IEnumerable<FunctionArgument>? value = item.Value as IEnumerable<FunctionArgument>;
                    if (value != null)
                    {
                        this.Calculate(value, ref nItems, context, ItemContext.InArray);
                    }
                    else
                    {
                        _CheckForAndHandleExcelError(item, context);
                        if (this.ShouldIgnore(item, context) == false && ShouldCount(item.Value, itemContext))
                        {
                            nItems++;
                        }
                    }
                }
            }
        }

        private static void _CheckForAndHandleExcelError(FunctionArgument arg, ParsingContext context)
        {
            //if (context.Scopes.Current.IsSubtotal)
            //{
            //    CheckForAndHandleExcelError(arg);
            //}
        }

        private static void _CheckForAndHandleExcelError(ICellInfo cell, ParsingContext context)
        {
            //if (context.Scopes.Current.IsSubtotal)
            //{
            //    CheckForAndHandleExcelError(cell);
            //}
        }

        private static bool ShouldCount(object value, ItemContext context)
        {
            switch (context)
            {
                case ItemContext.SingleArg:
                    return IsNumeric(value) || IsNumericString(value);
                case ItemContext.InRange:
                    return IsNumeric(value);
                case ItemContext.InArray:
                    return IsNumeric(value) || IsNumericString(value);
                default:
                    throw new ArgumentException("Unknown ItemContext:" + context.ToString());
            }
        }
    }
}
