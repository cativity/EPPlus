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

[FunctionMetadata(Category = ExcelFunctionCategory.MathAndTrig,
                  EPPlusVersion = "4",
                  Description = "Performs a specified calculation (e.g. the sum, product, average, etc.) for a supplied set of values")]
internal class Subtotal : ExcelFunction
{
    private Dictionary<int, HiddenValuesHandlingFunction> _functions = new Dictionary<int, HiddenValuesHandlingFunction>();

    public Subtotal() => this.Initialize();

    private void Initialize()
    {
        this._functions[1] = new Average();
        this._functions[2] = new Count();
        this._functions[3] = new CountA();
        this._functions[4] = new Max();
        this._functions[5] = new Min();
        this._functions[6] = new Product();
        this._functions[7] = new Stdev();
        this._functions[8] = new StdevP();
        this._functions[9] = new Sum();
        this._functions[10] = new Var();
        this._functions[11] = new VarP();

        this.AddHiddenValueHandlingFunction(new Average(), 101);
        this.AddHiddenValueHandlingFunction(new Count(), 102);
        this.AddHiddenValueHandlingFunction(new CountA(), 103);
        this.AddHiddenValueHandlingFunction(new Max(), 104);
        this.AddHiddenValueHandlingFunction(new Min(), 105);
        this.AddHiddenValueHandlingFunction(new Product(), 106);
        this.AddHiddenValueHandlingFunction(new Stdev(), 107);
        this.AddHiddenValueHandlingFunction(new StdevP(), 108);
        this.AddHiddenValueHandlingFunction(new Sum(), 109);
        this.AddHiddenValueHandlingFunction(new Var(), 110);
        this.AddHiddenValueHandlingFunction(new VarP(), 111);
    }

    private void AddHiddenValueHandlingFunction(HiddenValuesHandlingFunction func, int funcNum)
    {
        func.IgnoreHiddenValues = true;
        this._functions[funcNum] = func;
    }

    public override void BeforeInvoke(ParsingContext context)
    {
        base.BeforeInvoke(context);
        context.Scopes.Current.IsSubtotal = true;

        ulong cellId = context.ExcelDataProvider.GetCellId(context.Scopes.Current.Address.Worksheet,
                                                           context.Scopes.Current.Address.FromRow,
                                                           context.Scopes.Current.Address.FromCol);

        if (!context.SubtotalAddresses.Contains(cellId))
        {
            _ = context.SubtotalAddresses.Add(cellId);
        }
    }

    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        int funcNum = this.ArgToInt(arguments, 0);

        if (context.Scopes.Current.Parent != null && context.Scopes.Current.Parent.IsSubtotal)
        {
            return this.CreateResult(0d, DataType.Empty);
        }

        IEnumerable<FunctionArgument>? actualArgs = arguments.Skip(1);
        ExcelFunction function = this.GetFunctionByCalcType(funcNum);
        CompileResult? compileResult = function.Execute(actualArgs, context);
        compileResult.IsResultOfSubtotal = true;

        return compileResult;
    }

    private ExcelFunction GetFunctionByCalcType(int funcNum)
    {
        if (!this._functions.ContainsKey(funcNum))
        {
            ThrowExcelErrorValueException(eErrorType.Value);

            //throw new ArgumentException("Invalid funcNum " + funcNum + ", valid ranges are 1-11 and 101-111");
        }

        return this._functions[funcNum];
    }
}