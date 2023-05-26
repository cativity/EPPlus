using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class IgnoreCircularRefLookupCompiler : LookupFunctionCompiler
{
    public IgnoreCircularRefLookupCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
    {
    }

    public override CompileResult Compile(IEnumerable<Expression> children)
    {
        foreach(Expression? child in children)
        {
            child.IgnoreCircularReference = true;
        }
        return base.Compile(children);
    }
}