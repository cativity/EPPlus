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
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.UnrecognizedFunctionsPipeline;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

/// <summary>
/// Expression that handles execution of a function.
/// </summary>
internal class FunctionExpression : AtomicExpression
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="expression">should be the of the function</param>
    /// <param name="parsingContext"></param>
    /// <param name="isNegated">True if the numeric result of the function should be negated.</param>
    public FunctionExpression(string expression, ParsingContext parsingContext, bool isNegated)
        : base(expression)
    {
        this._parsingContext = parsingContext;
        this._functionCompilerFactory = new FunctionCompilerFactory(parsingContext.Configuration.FunctionRepository, parsingContext);
        this._isNegated = isNegated;
        _ = base.AddChild(new FunctionArgumentExpression(this));
    }

    private readonly ParsingContext _parsingContext;
    private readonly FunctionCompilerFactory _functionCompilerFactory;
    private readonly bool _isNegated;

    /// <summary>
    /// Compiles the expression
    /// </summary>
    /// <returns></returns>
    public override CompileResult Compile()
    {
        try
        {
            string? funcName = this.ExpressionString;

            // older versions of Excel (pre 2007) adds "_xlfn." in front of some function names for compatibility reasons.
            // EPPlus implements most of these functions, so we just remove this.
            if (funcName.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase))
            {
                funcName = funcName.Replace("_xlfn.", string.Empty);
            }

            ExcelFunction? function = this._parsingContext.Configuration.FunctionRepository.GetFunction(funcName);

            if (function == null)
            {
                // Handle unrecognized func name
                FunctionsPipeline? pipeline = new FunctionsPipeline(this._parsingContext, this.Children);
                function = pipeline.FindFunction(funcName);

                if (function == null)
                {
                    if (this._parsingContext.Debug)
                    {
                        this._parsingContext.Configuration.Logger.Log(this._parsingContext, string.Format("'{0}' is not a supported function", funcName));
                    }

                    return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);
                }
            }

            if (this._parsingContext.Debug)
            {
                this._parsingContext.Configuration.Logger.LogFunction(funcName);
            }

            FunctionCompiler? compiler = this._functionCompilerFactory.Create(function);
            CompileResult? result = compiler.Compile(this.HasChildren ? this.Children : Enumerable.Empty<Expression>());

            if (this._isNegated)
            {
                if (!result.IsNumeric)
                {
                    if (this._parsingContext.Debug)
                    {
                        string? msg = string.Format("Trying to negate a non-numeric value ({0}) in function '{1}'", result.Result, funcName);

                        this._parsingContext.Configuration.Logger.Log(this._parsingContext, msg);
                    }

                    return new CompileResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
                }

                return new CompileResult(result.ResultNumeric * -1, result.DataType);
            }

            return result;
        }
        catch (ExcelErrorValueException e)
        {
            if (this._parsingContext.Debug)
            {
                this._parsingContext.Configuration.Logger.Log(this._parsingContext, e);
            }

            return new CompileResult(e.ErrorValue, DataType.ExcelError);
        }
    }

    /// <summary>
    /// Adds a new <see cref="FunctionArgumentExpression"/> for the next child
    /// </summary>
    /// <returns></returns>
    public override Expression PrepareForNextChild() => base.AddChild(new FunctionArgumentExpression(this));

    /// <summary>
    /// Returns true if there are any existing children to this expression
    /// </summary>
    public override bool HasChildren => this.Children.Any() && this.Children.First().Children.Any();

    /// <summary>
    /// Adds a child expression
    /// </summary>
    /// <param name="child">The child expression to add</param>
    /// <returns></returns>
    public override Expression AddChild(Expression child)
    {
        _ = this.Children.Last().AddChild(child);

        return child;
    }
}