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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis;

[TestClass]
public class TokenFactoryTests
{
    private ITokenFactory _tokenFactory;
    private INameValueProvider _nameValueProvider;

    [TestInitialize]
    public void Setup()
    {
        ParsingContext? context = ParsingContext.Create();
        _ = A.Fake<ExcelDataProvider>();
        this._nameValueProvider = A.Fake<INameValueProvider>();
        this._tokenFactory = new TokenFactory(context.Configuration.FunctionRepository, this._nameValueProvider);
    }

    [TestCleanup]
    public void Cleanup()
    {
    }

    [TestMethod]
    public void ShouldCreateAStringToken()
    {
        string? input = "\"";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("\"", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.String));
    }

    [TestMethod]
    public void ShouldCreatePlusAsOperatorToken()
    {
        string? input = "+";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("+", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Operator));
    }

    [TestMethod]
    public void ShouldCreateMinusAsOperatorToken()
    {
        string? input = "-";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("-", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Operator));
    }

    [TestMethod]
    public void ShouldCreateMultiplyAsOperatorToken()
    {
        string? input = "*";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("*", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Operator));
    }

    [TestMethod]
    public void ShouldCreateDivideAsOperatorToken()
    {
        string? input = "/";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("/", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Operator));
    }

    [TestMethod]
    public void ShouldCreateEqualsAsOperatorToken()
    {
        string? input = "=";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("=", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Operator));
    }

    [TestMethod]
    public void ShouldCreateIntegerAsIntegerToken()
    {
        string? input = "23";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("23", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Integer));
    }

    [TestMethod]
    public void ShouldCreateBooleanAsBooleanToken()
    {
        string? input = "true";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("true", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Boolean));
    }

    [TestMethod]
    public void ShouldCreateDecimalAsDecimalToken()
    {
        string? input = "23.3";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);

        Assert.AreEqual("23.3", token.Value);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Decimal));
    }

    [TestMethod]
    public void CreateShouldReadFunctionsFromFuncRepository()
    {
        string? input = "Text";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.Function));
        Assert.AreEqual("Text", token.Value);
    }

    [TestMethod]
    public void CreateShouldCreateExcelAddressAsExcelAddressToken()
    {
        string? input = "A1";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.ExcelAddress));
        Assert.AreEqual("A1", token.Value);
    }

    [TestMethod]
    public void CreateShouldCreateExcelRangeAsExcelAddressToken()
    {
        string? input = "A1:B15";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.ExcelAddress));
        Assert.AreEqual("A1:B15", token.Value);
    }

    [TestMethod]
    public void CreateShouldCreateExcelRangeOnOtherSheetAsExcelAddressToken()
    {
        string? input = "ws!A1:B15";
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.ExcelAddress));
        Assert.AreEqual("ws!A1:B15", token.Value);
    }

    [TestMethod]
    public void CreateShouldCreateNamedValueAsExcelAddressToken()
    {
        string? input = "NamedValue";
        _ = A.CallTo(() => this._nameValueProvider.IsNamedValue("NamedValue", "")).Returns(true);
        _ = A.CallTo(() => this._nameValueProvider.IsNamedValue("NamedValue", null)).Returns(true);
        Token token = this._tokenFactory.Create(Enumerable.Empty<Token>(), input);
        Assert.IsTrue(token.TokenTypeIsSet(TokenType.NameValue));
        Assert.AreEqual("NamedValue", token.Value);
    }

    [TestMethod]
    public void TokenFactory_IsNumericTests()
    {
        List<Token>? tokens = new List<Token>();

        Token t = this._tokenFactory.Create(tokens, "1");
        Assert.IsTrue(t.TokenTypeIsSet(TokenType.Integer), "Failed to recognize integer");

        t = this._tokenFactory.Create(tokens, "1.01");
        Assert.IsTrue(t.TokenTypeIsSet(TokenType.Decimal), "Failed to recognize decimal");

        t = this._tokenFactory.Create(tokens, "1.01345E-05");
        Assert.IsTrue(t.TokenTypeIsSet(TokenType.Decimal), "Failed to recognize low exponential");

        t = this._tokenFactory.Create(tokens, "1.01345E+05");
        Assert.IsTrue(t.TokenTypeIsSet(TokenType.Decimal), "Failed to recognize high exponential");

        t = this._tokenFactory.Create(tokens, "ABC-E0");
        Assert.IsFalse(t.TokenTypeIsSet(TokenType.Integer | TokenType.Decimal), "Invalid low exponential number was still numeric 1");

        t = this._tokenFactory.Create(tokens, "ABC-E0");
        Assert.IsFalse(t.TokenTypeIsSet(TokenType.Integer | TokenType.Decimal), "Invalid high exponential number was still numeric 1");

        t = this._tokenFactory.Create(tokens, "E1");
        Assert.IsFalse(t.TokenTypeIsSet(TokenType.Integer | TokenType.Decimal), "Invalid exponential number was still numeric 2");
    }
}