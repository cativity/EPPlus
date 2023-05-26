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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class ExpressionGraphBuilder :IExpressionGraphBuilder
    {
        private readonly ExpressionGraph _graph = new ExpressionGraph();
        private readonly IExpressionFactory _expressionFactory;
        private readonly ParsingContext _parsingContext;
        private int _tokenIndex = 0;
        private int _nRangeOffsetTokens = 0;
        private RangeOffsetExpression _rangeOffsetExpression;
        private bool _negateNextExpression;

        public ExpressionGraphBuilder(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
            : this(new ExpressionFactory(excelDataProvider, parsingContext), parsingContext)
        {

        }

        public ExpressionGraphBuilder(IExpressionFactory expressionFactory, ParsingContext parsingContext)
        {
            this._expressionFactory = expressionFactory;
            this._parsingContext = parsingContext;
        }

        public ExpressionGraph Build(IEnumerable<Token> tokens)
        {
            this._tokenIndex = 0;
            this._graph.Reset();
            Token[]? tokensArr = tokens != null ? tokens.ToArray() : new Token[0];
            this.BuildUp(tokensArr, null);
            return this._graph;
        }

        private void BuildUp(Token[] tokens, Expression parent)
        {
            while (this._tokenIndex < tokens.Length)
            {
                Token token = tokens[this._tokenIndex];

                if (token.TokenTypeIsSet(TokenType.Operator) && OperatorsDict.Instance.TryGetValue(token.Value, out IOperator op))
                {
                    this.SetOperatorOnExpression(parent, op);
                }
                else if (token.TokenTypeIsSet(TokenType.RangeOffset))
                {
                    this.BuildRangeOffsetExpression(tokens, parent, token);
                }
                else if (token.TokenTypeIsSet(TokenType.Function))
                {
                    this.BuildFunctionExpression(tokens, parent, token.Value);
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningEnumerable))
                {
                    this._tokenIndex++;
                    this.BuildEnumerableExpression(tokens, parent);
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningParenthesis))
                {
                    this._tokenIndex++;
                    this.BuildGroupExpression(tokens, parent);
                }
                else if (token.TokenTypeIsSet(TokenType.ClosingParenthesis) || token.TokenTypeIsSet(TokenType.ClosingEnumerable))
                {
                    break;
                }
                else if(token.TokenTypeIsSet(TokenType.WorksheetName))
                {
                    StringBuilder? sb = new StringBuilder();
                    sb.Append(tokens[this._tokenIndex++].Value);
                    sb.Append(tokens[this._tokenIndex++].Value);
                    sb.Append(tokens[this._tokenIndex++].Value);
                    sb.Append(tokens[this._tokenIndex].Value);
                    Token t = new Token(sb.ToString(), TokenType.ExcelAddress);
                    this.CreateAndAppendExpression(ref parent, ref t);
                }
                else if (token.TokenTypeIsSet(TokenType.Negator))
                {
                    this._negateNextExpression = !this._negateNextExpression;
                }
                else if(token.TokenTypeIsSet(TokenType.Percent))
                {
                    this.SetOperatorOnExpression(parent, Operator.Percent);
                    if (parent == null)
                    {
                        this._graph.Add(ConstantExpressions.Percent);
                    }
                    else
                    {
                        parent.AddChild(ConstantExpressions.Percent);
                    }
                }
                else
                {
                    this.CreateAndAppendExpression(ref parent, ref token);
                }

                this._tokenIndex++;
            }
        }

        private void BuildEnumerableExpression(Token[] tokens, Expression parent)
        {
            if (parent == null)
            {
                this._graph.Add(new EnumerableExpression());
                this.BuildUp(tokens, this._graph.Current);
            }
            else
            {
                EnumerableExpression? enumerableExpression = new EnumerableExpression();
                parent.AddChild(enumerableExpression);
                this.BuildUp(tokens, enumerableExpression);
            }
        }

        private void CreateAndAppendExpression(ref Expression parent, ref Token token)
        {
            if (IsWaste(token))
            {
                return;
            }

            if (parent != null && 
                (token.TokenTypeIsSet(TokenType.Comma) || token.TokenTypeIsSet(TokenType.SemiColon)))
            {
                parent = parent.PrepareForNextChild();
                return;
            }
            if (this._negateNextExpression)
            {
                token = token.CloneWithNegatedValue(true);
                this._negateNextExpression = false;
            }
            Expression? expression = this._expressionFactory.Create(token);
            if (parent == null)
            {
                this._graph.Add(expression);
            }
            else
            {
                parent.AddChild(expression);
            }
        }

        private static bool IsWaste(Token token)
        {
            if (token.TokenTypeIsSet(TokenType.String) || token.TokenTypeIsSet(TokenType.Colon))
            {
                return true;
            }
            return false;
        }

        private void BuildRangeOffsetExpression(Token[] tokens, Expression parent, Token token)
        {
            if(this._nRangeOffsetTokens++ % 2 == 0)
            {
                this._rangeOffsetExpression = new RangeOffsetExpression(this._parsingContext);
                if(token.TokenTypeIsSet(TokenType.Function) && token.Value.ToLower() == "offset")
                {
                    this._rangeOffsetExpression.OffsetExpression1 = new FunctionExpression("offset", this._parsingContext, false);
                    this.HandleFunctionArguments(tokens, this._rangeOffsetExpression.OffsetExpression1);
                }
                else if(token.TokenTypeIsSet(TokenType.ExcelAddress))
                {
                    this._rangeOffsetExpression.AddressExpression2 = this._expressionFactory.Create(token) as ExcelAddressExpression;
                }
            }
            else
            {
                if (parent == null)
                {
                    this._graph.Add(this._rangeOffsetExpression);
                }
                else
                {
                    parent.AddChild(this._rangeOffsetExpression);
                }
                if (token.TokenTypeIsSet(TokenType.Function) && token.Value.ToLower() == "offset")
                {
                    this._rangeOffsetExpression.OffsetExpression2 = new FunctionExpression("offset", this._parsingContext, this._negateNextExpression);
                    this.HandleFunctionArguments(tokens, this._rangeOffsetExpression.OffsetExpression2);
                }
                else if (token.TokenTypeIsSet(TokenType.ExcelAddress))
                {
                    this._rangeOffsetExpression.AddressExpression2 = this._expressionFactory.Create(token) as ExcelAddressExpression;
                }
            }
        }

        private void BuildFunctionExpression(Token[] tokens, Expression parent, string funcName)
        {
            if (parent == null)
            {
                this._graph.Add(new FunctionExpression(funcName, this._parsingContext, this._negateNextExpression));
                this._negateNextExpression = false;
                this.HandleFunctionArguments(tokens, this._graph.Current);
            }
            else
            {
                FunctionExpression? func = new FunctionExpression(funcName, this._parsingContext, this._negateNextExpression);
                this._negateNextExpression = false;
                parent.AddChild(func);
                this.HandleFunctionArguments(tokens, func);
            }
        }

        private void HandleFunctionArguments(Token[] tokens, Expression function)
        {
            this._tokenIndex++;
            Token token = tokens.ElementAt(this._tokenIndex);
            if (!token.TokenTypeIsSet(TokenType.OpeningParenthesis))
            {
                throw new ExcelErrorValueException(eErrorType.Value);
            }

            this._tokenIndex++;
            this.BuildUp(tokens, function.Children.First());
        }

        private void BuildGroupExpression(Token[] tokens, Expression parent)
        {
            if (parent == null)
            {
                this._graph.Add(new GroupExpression(this._negateNextExpression));
                this._negateNextExpression = false;
                this.BuildUp(tokens, this._graph.Current);
            }
            else
            {
                if (parent.IsGroupedExpression || parent is FunctionArgumentExpression)
                {
                    GroupExpression? newGroupExpression = new GroupExpression(this._negateNextExpression);
                    this._negateNextExpression = false;
                    parent.AddChild(newGroupExpression);
                    this.BuildUp(tokens, newGroupExpression);
                }

                this.BuildUp(tokens, parent);
            }
        }

        private void SetOperatorOnExpression(Expression parent, IOperator op)
        {
            if (parent == null)
            {
                this._graph.Current.Operator = op;
            }
            else
            {
                Expression candidate;
                if (parent is FunctionArgumentExpression)
                {
                    candidate = parent.Children.Last();
                }
                else
                {
                    candidate = parent.Children.Last();
                    if (candidate is FunctionArgumentExpression)
                    {
                        candidate = candidate.Children.Last();
                    }
                }
                candidate.Operator = op;
            }
        }
    }
}
