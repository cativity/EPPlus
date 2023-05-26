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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.PostProcessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class TokenNaviagorTests
    {
        [TestMethod]
        public void ShouldNotHaveNextWhenOnlyOneToken()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);

            Assert.IsFalse(navigator.HasNext());
        }

        [TestMethod]
        public void ShouldHaveNextWhenMoreTokens()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);

            Assert.IsTrue(navigator.HasNext());
        }

        [TestMethod]
        public void ShouldNotHavePrevWheFirstToken()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);

            Assert.AreEqual(0, navigator.Index, "Index was not 0 but " + navigator.Index);
            Assert.IsFalse(navigator.HasPrev(), "HasPrev() was not false");
        }

        [TestMethod]
        public void IndexShouldIncreaseWhenMoveNext()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);
            navigator.MoveNext();

            Assert.AreEqual(1, navigator.Index, "Index was not 1 but " + navigator.Index);
        }

        [TestMethod]
        public void NextTokenShouldBeReturned()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);

            Assert.AreEqual("2", navigator.NextToken.Value);
        }

        [TestMethod]
        public void MoveToNextAndReturnPrevToken()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);
            navigator.MoveNext();

            Assert.AreEqual("1", navigator.PreviousToken.Value.Value);
        }

        [TestMethod]
        public void GetRelativeForward()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal),
                new Token("3", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);
            Token token = navigator.GetTokenAtRelativePosition(2);
            Assert.AreEqual("3", token.Value);
        }

        [TestMethod]
        public void NumberOfRemainingTokensShouldBeCorrect()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal),
                new Token("3", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);

            Assert.AreEqual(2, navigator.NbrOfRemainingTokens);
            navigator.MoveNext();
            Assert.AreEqual(1, navigator.NbrOfRemainingTokens);
        }

        [TestMethod]
        public void MoveIndexShouldSetNewPosition()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal),
                new Token("3", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);
            navigator.MoveIndex(1);
            Assert.AreEqual("2", navigator.CurrentToken.Value);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void ShouldThrowWhenIndexMovedOutOfRange()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal),
                new Token("2", TokenType.Decimal),
                new Token("3", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);
            navigator.MoveIndex(3);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ShouldThrowWhenGetPreviousOutOfRange()
        {
            List<Token>? tokens = new List<Token>
            {
                new Token("1", TokenType.Decimal)
            };
            TokenNavigator? navigator = new TokenNavigator(tokens);
            Token token = navigator.PreviousToken.Value;
        }
    }
}
