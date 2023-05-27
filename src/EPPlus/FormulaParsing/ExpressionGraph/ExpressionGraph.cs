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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExpressionGraph
{
    private List<Expression> _expressions = new List<Expression>();

    public IEnumerable<Expression> Expressions
    {
        get { return this._expressions; }
    }

    public Expression Current { get; private set; }

    public Expression Add(Expression expression)
    {
        this._expressions.Add(expression);

        if (this.Current != null)
        {
            this.Current.Next = expression;
            expression.Prev = this.Current;
        }

        this.Current = expression;

        return expression;
    }

    public void Reset()
    {
        this._expressions.Clear();
        this.Current = null;
    }

    public void Remove(Expression item)
    {
        if (item == this.Current)
        {
            this.Current = item.Prev ?? item.Next;
        }

        this._expressions.Remove(item);
    }
}