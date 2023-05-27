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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Utilities;
using IndexFunc = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Index;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class FunctionCompilerFactory
{
    private readonly Dictionary<Type, FunctionCompiler> _specialCompilers = new Dictionary<Type, FunctionCompiler>();
    private readonly ParsingContext _context;

    public FunctionCompilerFactory(FunctionRepository repository, ParsingContext context)
    {
        Require.That(context).Named("context").IsNotNull();
        this._context = context;
        this._specialCompilers.Add(typeof(If), new IfFunctionCompiler(repository.GetFunction("if"), context));
        this._specialCompilers.Add(typeof(CountIf), new CountIfFunctionCompiler(repository.GetFunction("countif"), context));
        this._specialCompilers.Add(typeof(SumIf), new SumIfCompiler(repository.GetFunction("sumif"), context));
        this._specialCompilers.Add(typeof(CountIfs), new CountIfsCompiler(repository.GetFunction("countifs"), context));
        this._specialCompilers.Add(typeof(IfError), new IfErrorFunctionCompiler(repository.GetFunction("iferror"), context));
        this._specialCompilers.Add(typeof(IfNa), new IfNaFunctionCompiler(repository.GetFunction("ifna"), context));
        this._specialCompilers.Add(typeof(Row), new IgnoreCircularRefLookupCompiler(repository.GetFunction("row"), context));
        this._specialCompilers.Add(typeof(Rows), new IgnoreCircularRefLookupCompiler(repository.GetFunction("rows"), context));
        this._specialCompilers.Add(typeof(Column), new IgnoreCircularRefLookupCompiler(repository.GetFunction("column"), context));
        this._specialCompilers.Add(typeof(Columns), new IgnoreCircularRefLookupCompiler(repository.GetFunction("columns"), context));
        this._specialCompilers.Add(typeof(IndexFunc), new IgnoreCircularRefLookupCompiler(repository.GetFunction("index"), context));

        foreach (Type? key in repository.CustomCompilers.Keys)
        {
            this._specialCompilers.Add(key, repository.CustomCompilers[key]);
        }
    }

    private FunctionCompiler GetCompilerByType(ExcelFunction function)
    {
        Type? funcType = function.GetType();

        if (this._specialCompilers.ContainsKey(funcType))
        {
            return this._specialCompilers[funcType];
        }
        else if (function.IsLookupFuction)
        {
            return new LookupFunctionCompiler(function, this._context);
        }
        else if (function.IsErrorHandlingFunction)
        {
            return new ErrorHandlingFunctionCompiler(function, this._context);
        }

        return new DefaultCompiler(function, this._context);
    }

    public virtual FunctionCompiler Create(ExcelFunction function)
    {
        return this.GetCompilerByType(function);
    }
}