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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of non-blanks in a supplied set of cells or values")]
    internal class CountA : HiddenValuesHandlingFunction
    {
        public CountA() : base()
        {
            this.IgnoreErrors = false;
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double nItems = 0d;
            this.Calculate(arguments, context, ref nItems);
            return this.CreateResult(nItems, DataType.Integer);
        }

        private void Calculate(IEnumerable<FunctionArgument> items, ParsingContext context, ref double nItems)
        {
            foreach (FunctionArgument? item in items)
            {
                IRangeInfo? cs = item.Value as IRangeInfo;
                if (cs != null)
                {
                    foreach (ICellInfo? c in cs)
                    {
                        _CheckForAndHandleExcelError(c, context);
                        if (!this.ShouldIgnore(c, context) && ShouldCount(c.Value))
                        {
                            nItems++;
                        }
                    }
                }
                else if (item.Value is IEnumerable<FunctionArgument>)
                {
                    this.Calculate((IEnumerable<FunctionArgument>)item.Value, context, ref nItems);
                }
                else
                {
                    _CheckForAndHandleExcelError(item, context);
                    if (!this.ShouldIgnore(item, context) && ShouldCount(item.Value))
                    {
                        nItems++;
                    }
                }

            }
        }

        private static void _CheckForAndHandleExcelError(FunctionArgument arg, ParsingContext context)
        {
            if (context.Scopes.Current.IsSubtotal)
            {
                CheckForAndHandleExcelError(arg);
            }
        }

        private static void _CheckForAndHandleExcelError(ICellInfo cell, ParsingContext context)
        {
            if (context.Scopes.Current.IsSubtotal)
            {
                CheckForAndHandleExcelError(cell);
            }
        }

        private static bool ShouldCount(object value)
        {
            return value != null;
        }
    }
}
