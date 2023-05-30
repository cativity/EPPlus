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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class IntegerExpression : AtomicExpression
{
    private double? _compiledValue;
    private bool _negate;

    public IntegerExpression(string expression)
        : this(expression, false)
    {
    }

    public IntegerExpression(string expression, bool negate)
        : base(expression) =>
        this._negate = negate;

    public IntegerExpression(double val)
        : base(val.ToString(CultureInfo.InvariantCulture)) =>
        this._compiledValue = Math.Floor(val);

    public override CompileResult Compile()
    {
        double result = this._compiledValue ?? double.Parse(this.ExpressionString, CultureInfo.InvariantCulture);
        result = this._negate ? result * -1 : result;

        return new CompileResult(result, DataType.Integer);
    }
}