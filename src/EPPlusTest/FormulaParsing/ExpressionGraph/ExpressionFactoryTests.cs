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
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.ExpressionGraph;

[TestClass]
public class ExpressionFactoryTests
{
    private IExpressionFactory _factory;
    private ParsingContext _parsingContext;

    [TestInitialize]
    public void Setup()
    {
        this._parsingContext = ParsingContext.Create();
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        this._factory = new ExpressionFactory(provider, this._parsingContext);
    }

    [TestMethod]
    public void ShouldReturnIntegerExpressionWhenTokenIsInteger()
    {
        Token token = new Token("2", TokenType.Integer);
        Expression? expression = this._factory.Create(token);
        Assert.IsInstanceOfType(expression, typeof(IntegerExpression));
    }

    [TestMethod]
    public void ShouldReturnBooleanExpressionWhenTokenIsBoolean()
    {
        Token token = new Token("true", TokenType.Boolean);
        Expression? expression = this._factory.Create(token);
        Assert.IsInstanceOfType(expression, typeof(BooleanExpression));
    }

    [TestMethod]
    public void ShouldReturnDecimalExpressionWhenTokenIsDecimal()
    {
        Token token = new Token("2.5", TokenType.Decimal);
        Expression? expression = this._factory.Create(token);
        Assert.IsInstanceOfType(expression, typeof(DecimalExpression));
    }

    [TestMethod]
    public void ShouldReturnExcelRangeExpressionWhenTokenIsExcelAddress()
    {
        Token token = new Token("A1", TokenType.ExcelAddress);
        Expression? expression = this._factory.Create(token);
        Assert.IsInstanceOfType(expression, typeof(ExcelAddressExpression));
    }

    [TestMethod]
    public void ShouldReturnNamedValueExpressionWhenTokenIsNamedValue()
    {
        Token token = new Token("NamedValue", TokenType.NameValue);
        Expression? expression = this._factory.Create(token);
        Assert.IsInstanceOfType(expression, typeof(NamedValueExpression));
    }
}