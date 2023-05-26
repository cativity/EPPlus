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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public abstract class Expression
{
    internal string ExpressionString { get; private set; }
    private readonly List<Expression> _children = new List<Expression>();
    public IEnumerable<Expression> Children { get { return this._children; } }
    public Expression Next { get; set; }
    public Expression Prev { get; set; }
    public IOperator Operator { get; set; }
    public abstract bool IsGroupedExpression { get; }
    /// <summary>
    /// If set to true, <see cref="ExcelAddressExpression"></see>s that has a circular reference to their cell will be ignored when compiled
    /// </summary>
    public virtual bool IgnoreCircularReference
    {
        get; set;
    }

    public Expression()
    {

    }

    public Expression(string expression)
    {
        this.ExpressionString = expression;
        this.Operator = null;
    }

    public virtual bool HasChildren
    {
        get { return this._children.Any(); }
    }

    public virtual Expression  PrepareForNextChild()
    {
        return this;
    }

    public virtual Expression AddChild(Expression child)
    {
        if (this._children.Any())
        {
            Expression? last = this._children.Last();
            child.Prev = last;
            last.Next = child;
        }

        this._children.Add(child);
        return child;
    }

    public virtual Expression MergeWithNext()
    {
        Expression? expression = this;
        if (this.Next != null && this.Operator != null)
        {
            CompileResult? result = this.Operator.Apply(this.Compile(), this.Next.Compile());
            expression = ExpressionConverter.Instance.FromCompileResult(result);
            if (expression is ExcelErrorExpression)
            {
                expression.Next = null;
                expression.Prev = null;
                return expression;
            }
            if (this.Next != null)
            {
                expression.Operator = this.Next.Operator;
            }
            else
            {
                expression.Operator = null;
            }
            expression.Next = this.Next.Next;
            if (expression.Next != null)
            {
                expression.Next.Prev = expression;
            }

            expression.Prev = this.Prev;
        }
        else
        {
            throw (new FormatException("Invalid formula syntax. Operator missing expression."));
        }
        if (this.Prev != null)
        {
            this.Prev.Next = expression;
        }            
        return expression;
    }

    public abstract CompileResult Compile();

}