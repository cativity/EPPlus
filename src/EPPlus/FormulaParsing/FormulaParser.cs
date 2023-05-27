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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;
using System.Diagnostics;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing;

/// <summary>
/// Entry class for the formula calulation engine of EPPlus.
/// </summary>
public class FormulaParser : IDisposable
{
    private readonly ParsingContext _parsingContext;
    private readonly ExcelDataProvider _excelDataProvider;

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="package">The package to calculate</param>
    public FormulaParser(ExcelPackage package)
        : this(new EpplusExcelDataProvider(package))
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="excelDataProvider">An instance of <see cref="ExcelDataProvider"/> which provides access to a workbook</param>
    internal FormulaParser(ExcelDataProvider excelDataProvider)
        : this(excelDataProvider, ParsingContext.Create())
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="excelDataProvider">An <see cref="ExcelDataProvider"></see></param>
    /// <param name="parsingContext">Parsing context</param>
    internal FormulaParser(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
    {
        parsingContext.Parser = this;
        parsingContext.ExcelDataProvider = excelDataProvider;
        parsingContext.NameValueProvider = new EpplusNameValueProvider(excelDataProvider);
        parsingContext.RangeAddressFactory = new RangeAddressFactory(excelDataProvider);
        this._parsingContext = parsingContext;
        this._excelDataProvider = excelDataProvider;

        this.Configure(configuration =>
        {
            configuration.SetLexer(new Lexer(this._parsingContext.Configuration.FunctionRepository, this._parsingContext.NameValueProvider))
                         .SetGraphBuilder(new ExpressionGraphBuilder(excelDataProvider, this._parsingContext))
                         .SetExpresionCompiler(new ExpressionCompiler())
                         .FunctionRepository.LoadModule(new BuiltInFunctions());
        });
    }

    /// <summary>
    /// This method enables configuration of the formula parser.
    /// </summary>
    /// <param name="configMethod">An instance of the </param>
    internal void Configure(Action<ParsingConfiguration> configMethod)
    {
        configMethod.Invoke(this._parsingContext.Configuration);
        this._lexer = this._parsingContext.Configuration.Lexer ?? this._lexer;
        this._graphBuilder = this._parsingContext.Configuration.GraphBuilder ?? this._graphBuilder;
        this._compiler = this._parsingContext.Configuration.ExpressionCompiler ?? this._compiler;
    }

    private ILexer _lexer;
    private IExpressionGraphBuilder _graphBuilder;
    private IExpressionCompiler _compiler;

    internal ILexer Lexer
    {
        get { return this._lexer; }
    }

    internal IEnumerable<string> FunctionNames
    {
        get { return this._parsingContext.Configuration.FunctionRepository.FunctionNames; }
    }

    /// <summary>
    /// Contains information about filters on a workbook's worksheets.
    /// </summary>
    internal FilterInfo FilterInfo { get; private set; }

    internal virtual object Parse(string formula, RangeAddress rangeAddress)
    {
        using ParsingScope? scope = this._parsingContext.Scopes.NewScope(rangeAddress);
        IEnumerable<Token>? tokens = this._lexer.Tokenize(formula);
        ExpressionGraph.ExpressionGraph? graph = this._graphBuilder.Build(tokens);

        if (graph.Expressions.Count() == 0)
        {
            return null;
        }

        return this._compiler.Compile(graph.Expressions).Result;
    }

    internal virtual object Parse(IEnumerable<Token> tokens, string worksheet, string address)
    {
        RangeAddress? rangeAddress = this._parsingContext.RangeAddressFactory.Create(address);
        using ParsingScope? scope = this._parsingContext.Scopes.NewScope(rangeAddress);
        ExpressionGraph.ExpressionGraph? graph = this._graphBuilder.Build(tokens);

        if (graph.Expressions.Count() == 0)
        {
            return null;
        }

        return this._compiler.Compile(graph.Expressions).Result;
    }

    internal virtual object ParseCell(IEnumerable<Token> tokens, string worksheet, int row, int column)
    {
        RangeAddress? rangeAddress = this._parsingContext.RangeAddressFactory.Create(worksheet, column, row);
        using ParsingScope? scope = this._parsingContext.Scopes.NewScope(rangeAddress);

        //    _parsingContext.Dependencies.AddFormulaScope(scope);
        ExpressionGraph.ExpressionGraph? graph = this._graphBuilder.Build(tokens);

        if (graph.Expressions.Count() == 0)
        {
            return 0d;
        }

        try
        {
            CompileResult? compileResult = this._compiler.Compile(graph.Expressions);

            // quick solution for the fact that an excelrange can be returned.
            IRangeInfo? rangeInfo = compileResult.Result as IRangeInfo;

            if (rangeInfo == null)
            {
                return compileResult.Result ?? 0d;
            }
            else
            {
                if (rangeInfo.IsEmpty)
                {
                    return 0d;
                }

                if (!rangeInfo.IsMulti)
                {
                    return rangeInfo.First().Value ?? 0d;
                }

                // ok to return multicell if it is a workbook scoped name.
                if (string.IsNullOrEmpty(worksheet))
                {
                    return rangeInfo;
                }

                if (this._parsingContext.Debug)
                {
                    string? msg = string.Format("A range with multiple cell was returned at row {0}, column {1}", row, column);

                    this._parsingContext.Configuration.Logger.Log(this._parsingContext, msg);
                }

                return ExcelErrorValue.Create(eErrorType.Value);
            }
        }
        catch (ExcelErrorValueException ex)
        {
            if (this._parsingContext.Debug)
            {
                this._parsingContext.Configuration.Logger.Log(this._parsingContext, ex);
            }

            return ex.ErrorValue;
        }
    }

    /// <summary>
    /// Parses a formula at a specific address
    /// </summary>
    /// <param name="formula">A string containing the formula</param>
    /// <param name="address">Address of the formula</param>
    /// <returns></returns>
    public virtual object Parse(string formula, string address)
    {
        return this.Parse(formula, this._parsingContext.RangeAddressFactory.Create(address));
    }

    /// <summary>
    /// Parses a formula
    /// </summary>
    /// <param name="formula">A string containing the formula</param>
    /// <returns>The result of the calculation</returns>
    public virtual object Parse(string formula)
    {
        return this.Parse(formula, RangeAddress.Empty);
    }

    /// <summary>
    /// Parses a formula in a specific location
    /// </summary>
    /// <param name="address">address of the cell to calculate</param>
    /// <returns>The result of the calculation</returns>
    public virtual object ParseAt(string address)
    {
        Require.That(address).Named("address").IsNotNullOrEmpty();
        RangeAddress? rangeAddress = this._parsingContext.RangeAddressFactory.Create(address);

        return this.ParseAt(rangeAddress.Worksheet, rangeAddress.FromRow, rangeAddress.FromCol);
    }

    /// <summary>
    /// Parses a formula in a specific location
    /// </summary>
    /// <param name="worksheetName">Name of the worksheet</param>
    /// <param name="row">Row in the worksheet</param>
    /// <param name="col">Column in the worksheet</param>
    /// <returns>The result of the calculation</returns>
    public virtual object ParseAt(string worksheetName, int row, int col)
    {
        string? f = this._excelDataProvider.GetRangeFormula(worksheetName, row, col);

        if (string.IsNullOrEmpty(f))
        {
            return this._excelDataProvider.GetRangeValue(worksheetName, row, col);
        }
        else
        {
            return this.Parse(f, this._parsingContext.RangeAddressFactory.Create(worksheetName, col, row));
        }

        //var dataItem = _excelDataProvider.GetRangeValues(address).FirstOrDefault();
        //if (dataItem == null /*|| (dataItem.Value == null && dataItem.Formula == null)*/) return null;
        //if (!string.IsNullOrEmpty(dataItem.Formula))
        //{
        //    return Parse(dataItem.Formula, _parsingContext.RangeAddressFactory.Create(address));
        //}
        //return Parse(dataItem.Value.ToString(), _parsingContext.RangeAddressFactory.Create(address));
    }

    internal void InitNewCalc(FilterInfo filterInfo)
    {
        this.FilterInfo = filterInfo;

        if (this._excelDataProvider != null)
        {
            this._excelDataProvider.Reset();
        }
    }

    /// <summary>
    /// An <see cref="IFormulaParserLogger"/> for logging during calculation
    /// </summary>
    public IFormulaParserLogger Logger
    {
        get { return this._parsingContext.Configuration.Logger; }
    }

    /// <summary>
    /// Implementation of <see cref="IDisposable"></see>
    /// </summary>
    public void Dispose()
    {
        if (this._parsingContext.Debug)
        {
            this._parsingContext.Configuration.Logger.Dispose();
        }
    }
}