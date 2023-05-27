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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class IfErrorFunctionCompiler : FunctionCompiler
{
    public IfErrorFunctionCompiler(ExcelFunction function, ParsingContext context)
        : base(function, context)
    {
        Require.That(function).Named("function").IsNotNull();
    }

    public override CompileResult Compile(IEnumerable<Expression> children)
    {
        if (children.Count() != 2)
        {
            throw new ExcelErrorValueException(eErrorType.Value);
        }

        List<FunctionArgument>? args = new List<FunctionArgument>();
        this.Function.BeforeInvoke(this.Context);
        Expression? firstChild = children.First();
        Expression? lastChild = children.ElementAt(1);

        try
        {
            CompileResult? result = firstChild.Compile();

            if (result.DataType == DataType.ExcelError)
            {
                args.Add(new FunctionArgument(lastChild.Compile().Result));
            }
            else
            {
                args.Add(new FunctionArgument(result.Result));
            }
        }
        catch (ExcelErrorValueException)
        {
            args.Add(new FunctionArgument(lastChild.Compile().Result));
        }

        return this.Function.Execute(args, this.Context);
    }
}