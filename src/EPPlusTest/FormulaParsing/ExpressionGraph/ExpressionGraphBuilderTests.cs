/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExpressionGraphBuilderTests
    {
        private IExpressionGraphBuilder _graphBuilder;
        private ExcelDataProvider _excelDataProvider;

        [TestInitialize]
        public void Setup()
        {
            this._excelDataProvider = A.Fake<ExcelDataProvider>();
            ParsingContext? parsingContext = ParsingContext.Create();
            this._graphBuilder = new ExpressionGraphBuilder(this._excelDataProvider, parsingContext);
        }

        [TestCleanup]
        public void Cleanup()
        {

        }

        [TestMethod]
        public void BuildShouldNotUseStringIdentifyersWhenBuildingStringExpression()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc", TokenType.StringContent),
                new Token("'", TokenType.String)
            };

            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.Count());
        }

        [TestMethod]
        public void BuildShouldNotEvaluateExpressionsWithinAString()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("1 + 2", TokenType.StringContent),
                new Token("'", TokenType.String)
            };

            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.AreEqual("1 + 2", result.Expressions.First().Compile().Result);
        }

        [TestMethod]
        public void BuildShouldSetOperatorOnGroupExpressionCorrectly()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.AreEqual(Operator.Multiply.Operator, result.Expressions.First().Operator.Operator);

        }

        [TestMethod]
        public void BuildShouldSetChildrenOnGroupExpression()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.IsInstanceOfType(result.Expressions.First(), typeof(GroupExpression));
            Assert.AreEqual(2, result.Expressions.First().Children.Count());
        }

        [TestMethod]
        public void BuildShouldSetNextOnGroupedExpression()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.IsNotNull(result.Expressions.First().Next);
            Assert.IsInstanceOfType(result.Expressions.First().Next, typeof(IntegerExpression));

        }

        [TestMethod]
        public void BuildShouldBuildFunctionExpressionIfFirstTokenIsFunction()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
            };
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.Count());
            Assert.IsInstanceOfType(result.Expressions.First(), typeof(FunctionExpression));
        }

        [TestMethod]
        public void BuildShouldSetChildrenOnFunctionExpression()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis)
            };
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.First().Children.Count());
            Assert.IsInstanceOfType(result.Expressions.First().Children.First(), typeof(GroupExpression));
            Assert.IsInstanceOfType(result.Expressions.First().Children.First().Children.First(), typeof(IntegerExpression));
            Assert.AreEqual(2d, result.Expressions.First().Children.First().Compile().Result);
        }

        [TestMethod]
        public void BuildShouldAddOperatorToFunctionExpression()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("&", TokenType.Operator),
                new Token("A", TokenType.StringContent)
            };
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.First().Children.Count());
            Assert.AreEqual(2, result.Expressions.Count());
        }

        [TestMethod]
        public void BuildShouldAddCommaSeparatedFunctionArgumentsAsChildrenToFunctionExpression()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("Text", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token(",", TokenType.Comma),
                new Token("3", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("&", TokenType.Operator),
                new Token("A", TokenType.StringContent)
            };

            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.AreEqual(2, result.Expressions.First().Children.Count());
        }

        [TestMethod]
        public void BuildShouldCreateASingleExpressionOutOfANegatorAndANumericToken()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("-", TokenType.Negator),
                new Token("2", TokenType.Integer),
            };

            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.Count());
            Assert.AreEqual(-2d, result.Expressions.First().Compile().Result);
        }

        [TestMethod]
        public void BuildShouldHandleEnumerableTokens()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("Text", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("{", TokenType.OpeningEnumerable),
                new Token("2", TokenType.Integer),
                new Token(",", TokenType.Comma),
                new Token("3", TokenType.Integer),
                new Token("}", TokenType.ClosingEnumerable),
                new Token(")", TokenType.ClosingParenthesis)
            };

            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);
            Expression? funcArgExpression = result.Expressions.First().Children.First();
            Assert.IsInstanceOfType(funcArgExpression, typeof(FunctionArgumentExpression));

            Expression? enumerableExpression = funcArgExpression.Children.First();

            Assert.IsInstanceOfType(enumerableExpression, typeof(EnumerableExpression));
            Assert.AreEqual(2, enumerableExpression.Children.Count(), "Enumerable.Count was not 2");
        }

        [TestMethod]
        public void ShouldHandleInnerFunctionCall2()
        {
            ParsingContext? ctx = ParsingContext.Create();
            const string formula = "IF(3>2;\"Yes\";\"No\")";
            SourceCodeTokenizer? tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
            IEnumerable<Token>? tokens = tokenizer.Tokenize(formula);
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? expression = this._graphBuilder.Build(tokens);
            Assert.AreEqual(1, expression.Expressions.Count());

            ExpressionCompiler? compiler = new ExpressionCompiler(new ExpressionConverter(), new CompileStrategyFactory());
            CompileResult? result = compiler.Compile(expression.Expressions);
            Assert.AreEqual("Yes", result.Result);
        }

        [TestMethod]
        public void ShouldHandleInnerFunctionCall3()
        {
            ParsingContext? ctx = ParsingContext.Create();
            const string formula = "IF(I10>=0;IF(O10>I10;((O10-I10)*$B10)/$C$27;IF(O10<0;(O10*$B10)/$C$27;\"\"));IF(O10<0;((O10-I10)*$B10)/$C$27;IF(O10>0;(O10*$B10)/$C$27;)))";
            SourceCodeTokenizer? tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
            IEnumerable<Token>? tokens = tokenizer.Tokenize(formula);
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? expression = this._graphBuilder.Build(tokens);
            Assert.AreEqual(1, expression.Expressions.Count());
            Expression? exp1 = expression.Expressions.First();
            Assert.AreEqual(3, exp1.Children.Count());
        }
        [TestMethod, Ignore]
        public void RemoveDuplicateOperators1()
        {
            ParsingContext? ctx = ParsingContext.Create();
            const string formula = "++1--2++-3+-1----3-+2";
            // the formula above equals 1+2-3-1+3+2
            SourceCodeTokenizer? tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
            List<Token>? tokens = tokenizer.Tokenize(formula).ToList();
            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? expression = this._graphBuilder.Build(tokens);
            Assert.AreEqual(11, tokens.Count);
            Assert.AreEqual("+", tokens[1].Value);
            Assert.AreEqual("-", tokens[3].Value);
            Assert.AreEqual("-", tokens[5].Value);
            Assert.AreEqual("+", tokens[7].Value);
            Assert.AreEqual("-", tokens[9].Value);
        }
        [TestMethod, Ignore]
        public void RemoveDuplicateOperators2()
        {
            ParsingContext? ctx = ParsingContext.Create();
            const string formula = "++-1--(---2)++-3+-1----3-+2";
            SourceCodeTokenizer? tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
            List<Token>? tokens = tokenizer.Tokenize(formula).ToList();
        }

        [TestMethod]
        public void BuildExcelAddressExpressionSimple()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("A1", TokenType.ExcelAddress)
            };

            OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph? result = this._graphBuilder.Build(tokens);
            Assert.IsInstanceOfType(result.Expressions.First(), typeof(ExcelAddressExpression));
        }
    }
}
