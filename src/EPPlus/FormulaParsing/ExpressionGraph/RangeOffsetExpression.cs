﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

/// <summary>
/// This Expression handles addresses where the OFFSET function is a part of the range, i.e. OFFSET(..):A1 or OFFSET(..):OFFSET(..)
/// </summary>
internal class RangeOffsetExpression : Expression
{
    public RangeOffsetExpression(ParsingContext context) => this._parsingContext = context;

    /// <summary>
    /// The first part of the range, should be an OFFSET call
    /// </summary>
    public FunctionExpression OffsetExpression1 { get; set; }

    /// <summary>
    /// The second part of the range, should be an OFFSET call
    /// </summary>
    public FunctionExpression OffsetExpression2 { get; set; }

    /// <summary>
    /// The second part of the range, should be an Excel address
    /// </summary>
    public ExcelAddressExpression AddressExpression2 { get; set; }

    private readonly ParsingContext _parsingContext;

    public override bool IsGroupedExpression => false;

    public override CompileResult Compile()
    {
        IRangeInfo? offsetRange1 = this.OffsetExpression1.Compile().Result as IRangeInfo;
        RangeOffset? rangeOffset = new RangeOffset { StartRange = offsetRange1 };

        if (this.AddressExpression2 != null)
        {
            ParsingScope? c = this._parsingContext.Scopes.Current;

            IRangeInfo? resultRange =
                this._parsingContext.ExcelDataProvider.GetRange(c.Address.Worksheet,
                                                                c.Address.FromRow,
                                                                c.Address.FromCol,
                                                                this.AddressExpression2.ExpressionString);

            rangeOffset.EndRange = resultRange;
        }
        else
        {
            object? offsetRange2 = this.OffsetExpression2.Compile().Result;
            rangeOffset.EndRange = offsetRange2 as IRangeInfo;
        }

        return new CompileResult(rangeOffset.Execute(new FunctionArgument[] { }, this._parsingContext).Result, DataType.Enumerable);
    }
}