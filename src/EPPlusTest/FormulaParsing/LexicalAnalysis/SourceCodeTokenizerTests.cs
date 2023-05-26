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
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class SourceCodeTokenizerTests
    {
        private SourceCodeTokenizer _tokenizer;

        [TestInitialize]
        public void Setup()
        {
            ParsingContext? context = ParsingContext.Create();
            this._tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, NameValueProvider.Empty);
        }

        [TestCleanup]
        public void Cleanup()
        {
        }

        [TestMethod]
        public void ShouldCreateTokensForStringCorrectly()
        {
            string? input = "\"abc123\"";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.IsTrue(tokens.First().TokenTypeIsSet(TokenType.String));
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.StringContent));
            Assert.IsTrue(tokens.Last().TokenTypeIsSet(TokenType.String));
        }

        [TestMethod]
        public void ShouldTokenizeStringCorrectly()
        {
            string? input = "\"ab(c)d\"";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
        }

        [TestMethod]
        public void ShouldHandleWhitespaceCorrectly()
        {
            string? input = @"""          """;
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);
            Assert.AreEqual(3, tokens.Count());
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.StringContent));
            Assert.AreEqual(10, tokens.ElementAt(1).Value.Length);
        }

        [TestMethod]
        public void ShouldCreateTokensForFunctionCorrectly()
        {
            string? input = "Text(2)";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);

            Assert.AreEqual(4, tokens.Count());
            Assert.IsTrue(tokens.First().TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens.ElementAt(2).TokenTypeIsSet(TokenType.Integer));
            Assert.AreEqual("2", tokens.ElementAt(2).Value);
            Assert.IsTrue(tokens.Last().TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void ShouldHandleMultipleCharOperatorCorrectly()
        {
            string? input = "1 <= 2";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual("<=", tokens.ElementAt(1).Value);
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.Operator));
        }

        [TestMethod]
        public void ShouldCreateTokensForEnumerableCorrectly()
        {
            string? input = "Text({1;2})";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();

            Assert.AreEqual(8, tokens.Length);
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.OpeningEnumerable));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.ClosingEnumerable));
        }

        [TestMethod]
        public void ShouldCreateTokensWithStringForEnumerableCorrectly()
        {
            string? input = "{\"1\",\"2\"}";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();

            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.OpeningEnumerable));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.String));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.StringContent));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingEnumerable));
        }

        [TestMethod]
        public void ShouldCreateTokensForExcelAddressCorrectly()
        {
            string? input = "Text(A1)";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);

            Assert.IsTrue(tokens.ElementAt(2).TokenTypeIsSet(TokenType.ExcelAddress));
        }

        [TestMethod]
        public void ShouldCreateTokenForPercentAfterDecimal()
        {
            string? input = "1,23%";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);
            Assert.IsTrue(tokens.Last().TokenTypeIsSet(TokenType.Percent));
        }

        [TestMethod]
        public void ShouldIgnoreTwoSubsequentStringIdentifyers()
        {
            string? input = "\"hello\"\"world\"";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);
            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual("hello\"world", tokens.ElementAt(1).Value);
        }

        [TestMethod]
        public void ShouldIgnoreTwoSubsequentStringIdentifyers2()
        {
            //using (var pck = new ExcelPackage(new FileInfo("c:\\temp\\QuoteIssue2.xlsx")))
            //{
            //    pck.Workbook.Worksheets.First().Calculate();
            //}
            string? input = "\"\"\"\"\"\"";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.StringContent));
        }

        [TestMethod]
        public void TokenizerShouldIgnoreOperatorInString()
        {
            string? input = "\"*\"";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);
            Assert.IsTrue(tokens.ElementAt(1).TokenTypeIsSet(TokenType.StringContent));
        }

        [TestMethod]
        public void TokenizerShouldHandleWorksheetNameWithMinus()
        {
            string? input = "'A-B'!A1";
            IEnumerable<Token>? tokens = this._tokenizer.Tokenize(input);
            Assert.AreEqual(1, tokens.Count());
            Assert.IsTrue(tokens.ElementAt(0).TokenTypeIsSet(TokenType.ExcelAddress));
        }

        [TestMethod]
        public void TestBug9_12_14()
        {
            //(( W60 -(- W63 )-( W29 + W30 + W31 ))/( W23 + W28 + W42 - W51 )* W4 )
            using ExcelPackage? pck = new ExcelPackage();
            ExcelWorksheet? ws1 = pck.Workbook.Worksheets.Add("test");
            for (int x = 1; x <= 10; x++)
            {
                ws1.Cells[x, 1].Value = x;
            }

            ws1.Cells["A11"].Formula = "(( A1 -(- A2 )-( A3 + A4 + A5 ))/( A6 + A7 + A8 - A9 )* A5 )";
            //ws1.Cells["A11"].Formula = "(-A2 + 1 )";
            ws1.Calculate();
            object? result = ws1.Cells["A11"].Value;
            Assert.AreEqual(-3.75, result);
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSign()
        {
            string? input = @"+3-3";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(3, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeIdentifiesDoubleNegator()
        {
            string? input = @"--3-3";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(5, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeHandlesPositiveNegator()
        {
            string? input = @"+-3-3";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(4, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositive()
        {
            string? input = @"-+3-3";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(4, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSignFromFirstFunctionArgument()
        {
            string? input = @"SUM(+3-3,5)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Length);

            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingPlusSignFromSecondFunctionArgument()
        {
            string? input = @"SUM(5,+3-3)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(8, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingDoubleNegatorFromFirstFunctionArgument()
        {
            string? input = @"SUM(--3-3,5)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(10, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[9].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeStripsLeadingDoubleNegatorFromSecondFunctionArgument()
        {
            string? input = @"SUM(5,--3-3)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(10, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function), "TokenType was not function");
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis), "TokenType was not OpeningParenthesis");
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer), "TokenType was not Integer 2");
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Negator), "TokenType was not negator 4");
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Negator), "TokenType was not negator 5");
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[9].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesPositiveNegatorAsFirstFunctionArgument()
        {
            string? input = @"SUM(+-3-3,5)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositiveAsFirstFunctionArgument()
        {
            string? input = @"SUM(-+3-3,5)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesPositiveNegatorAsSecondFunctionArgument()
        {
            string? input = @"SUM(5,+-3-3)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }

        [TestMethod]
        public void TokenizeHandlesNegatorPositiveAsSecondFunctionArgument()
        {
            string? input = @"SUM(5,-+3-3)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(9, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.Function));
            Assert.IsTrue(tokens[1].TokenTypeIsSet(TokenType.OpeningParenthesis));
            Assert.IsTrue(tokens[2].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[3].TokenTypeIsSet(TokenType.Comma));
            Assert.IsTrue(tokens[4].TokenTypeIsSet(TokenType.Negator));
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[6].TokenTypeIsSet(TokenType.Operator));
            Assert.IsTrue(tokens[7].TokenTypeIsSet(TokenType.Integer));
            Assert.IsTrue(tokens[8].TokenTypeIsSet(TokenType.ClosingParenthesis));
        }
        [TestMethod]
        public void TokenizeWorksheetName()
        {
            string? input = @"sheetname!name";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(1, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.NameValue));
        }

        [TestMethod]
        public void TokenizeWorksheetNameWithQuotes()
        {
            string? input = @"'sheetname'!name";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(1, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorksheetName()
        {
            string? input = @"[0]sheetname!name";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(1, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.NameValue));
        }

        [TestMethod]
        public void TokenizeExternalWorksheetNameWithQuotes()
        {
            string? input = @"[3]'sheetname'!name";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(1, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorkbookName()
        {
            string? input = @"[0]!name";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(1, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.NameValue));
        }
        [TestMethod]
        public void TokenizeExternalWorkbookInvalidRef()
        {
            string? input = @"[0]#Ref!";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(1, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.InvalidReference));
        }
        [TestMethod]
        public void TokenizeExternalWorksheetInvalidRef()
        {
            string? input = @"[0]Sheet1!#Ref!";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(1, tokens.Length);
            Assert.IsTrue(tokens[0].TokenTypeIsSet(TokenType.InvalidReference));
        }

        [TestMethod]
        public void TokenizeShouldHandleWorksheetNameWithSingleQuote()
        {
            string? input = @"=VLOOKUP(J7;'Sheet 1''21'!$Q$4:$R$28;2;0)";
            Token[]? tokens = this._tokenizer.Tokenize(input).ToArray();
            Assert.AreEqual(11, tokens.Length);
            Assert.IsTrue(tokens[5].TokenTypeIsSet(TokenType.ExcelAddress));
            Assert.AreEqual("'Sheet 1''21'!$Q$4:$R$28", tokens[5].Value);
        }
    }
}
