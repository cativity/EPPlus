﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExcelRangeExpression : Expression
{
    public ExcelRangeExpression(IRangeInfo rangeInfo) => this._rangeInfo = rangeInfo;

    private readonly IRangeInfo _rangeInfo;

    public override bool IsGroupedExpression => false;

    public override CompileResult Compile() => new(this._rangeInfo, DataType.Enumerable);
}