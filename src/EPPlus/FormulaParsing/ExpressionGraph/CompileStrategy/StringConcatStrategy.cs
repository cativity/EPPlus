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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

public class StringConcatStrategy : CompileStrategy
{
    public StringConcatStrategy(Expression expression)
        : base(expression)
    {
           
    }

    public override Expression Compile()
    {
        Expression? newExp = this._expression is ExcelAddressExpression ? this._expression : ExpressionConverter.Instance.ToStringExpression(this._expression);
        newExp.Prev = this._expression.Prev;
        newExp.Next = this._expression.Next;
        if (this._expression.Prev != null)
        {
            this._expression.Prev.Next = newExp;
        }
        if (this._expression.Next != null)
        {
            this._expression.Next.Prev = newExp;
        }
        return newExp.MergeWithNext();
    }
}