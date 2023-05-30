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

//using System.Diagnostics.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest;

[TestClass]
public class ExpressionEvaluatorTests
{
    private ExpressionEvaluator _evaluator;

    [TestInitialize]
    public void Setup() => this._evaluator = new ExpressionEvaluator();

    #region Numeric Expression Tests

    [TestMethod]
    public void EvaluateShouldReturnTrueIfOperandsAreEqual()
    {
        bool result = this._evaluator.Evaluate("1", "1");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateShouldReturnTrueIfOperandsAreMatchingButDifferentTypes()
    {
        bool result = this._evaluator.Evaluate(1d, "1");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateShouldEvaluateOperator()
    {
        bool result = this._evaluator.Evaluate(1d, "<2");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateShouldEvaluateNumericString()
    {
        bool result = this._evaluator.Evaluate("1", ">0");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateShouldHandleBooleanArg()
    {
        bool result = this._evaluator.Evaluate(true, "TRUE");
        Assert.IsTrue(result);
    }

    [TestMethod, ExpectedException(typeof(ArgumentException))]
    public void EvaluateShouldThrowIfOperatorIsNotBoolean() => _ = this._evaluator.Evaluate(1d, "+1");

    [TestMethod]
    public void EvaluateShouldEvaluateToGreaterThanMinusOne()
    {
        bool result = this._evaluator.Evaluate(1d, "<>-1");
        Assert.IsTrue(result);
    }

    #endregion

    #region Date tests

    [TestMethod]
    public void EvaluateShouldHandleDateArg()
    {
#if (!Core)
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
#endif
        bool result = this._evaluator.Evaluate(new DateTime(2016, 6, 28), "2016-06-28");
        Assert.IsTrue(result);
#if (!Core)
            Thread.CurrentThread.CurrentCulture = ci;
#endif
    }

    [TestMethod]
    public void EvaluateShouldHandleDateArgWithOperator()
    {
#if (!Core)
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
#endif
        bool result = this._evaluator.Evaluate(new DateTime(2016, 6, 28), ">2016-06-27");
        Assert.IsTrue(result);
#if (!Core)
            Thread.CurrentThread.CurrentCulture = ci;
#endif
    }

    #endregion

    #region Blank Expression Tests

    [TestMethod]
    public void EvaluateBlankExpressionEqualsNull()
    {
        bool result = this._evaluator.Evaluate(null, "");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateBlankExpressionEqualsEmptyString()
    {
        bool result = this._evaluator.Evaluate(string.Empty, "");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateBlankExpressionEqualsZero()
    {
        bool result = this._evaluator.Evaluate(0d, "");
        Assert.IsFalse(result);
    }

    #endregion

    #region Quotes Expression Tests

    [TestMethod]
    public void EvaluateQuotesExpressionEqualsNull()
    {
        bool result = this._evaluator.Evaluate(null, "\"\"");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateQuotesExpressionEqualsZero()
    {
        bool result = this._evaluator.Evaluate(0d, "\"\"");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateQuotesExpressionEqualsCharacter()
    {
        bool result = this._evaluator.Evaluate("a", "\"\"");
        Assert.IsFalse(result);
    }

    #endregion

    #region NotEqualToZero Expression Tests

    [TestMethod]
    public void EvaluateNotEqualToZeroExpressionEqualsNull()
    {
        bool result = this._evaluator.Evaluate(null, "<>0");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateNotEqualToZeroExpressionEqualsEmptyString()
    {
        bool result = this._evaluator.Evaluate(string.Empty, "<>0");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateNotEqualToZeroExpressionEqualsCharacter()
    {
        bool result = this._evaluator.Evaluate("a", "<>0");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateNotEqualToZeroExpressionEqualsNonZero()
    {
        bool result = this._evaluator.Evaluate(1d, "<>0");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateNotEqualToZeroExpressionEqualsZero()
    {
        bool result = this._evaluator.Evaluate(0d, "<>0");
        Assert.IsFalse(result);
    }

    #endregion

    #region NotEqualToBlank Expression Tests

    [TestMethod]
    public void EvaluateNotEqualToBlankExpressionEqualsNull()
    {
        bool result = this._evaluator.Evaluate(null, "<>");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateNotEqualToBlankExpressionEqualsEmptyString()
    {
        bool result = this._evaluator.Evaluate(string.Empty, "<>");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateNotEqualToBlankExpressionEqualsCharacter()
    {
        bool result = this._evaluator.Evaluate("a", "<>");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateNotEqualToBlankExpressionEqualsNonZero()
    {
        bool result = this._evaluator.Evaluate(1d, "<>");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateNotEqualToBlankExpressionEqualsZero()
    {
        bool result = this._evaluator.Evaluate(0d, "<>");
        Assert.IsTrue(result);
    }

    #endregion

    #region Character Expression Tests

    [TestMethod]
    public void EvaluateCharacterExpressionEqualNull()
    {
        bool result = this._evaluator.Evaluate(null, "a");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateCharacterExpressionEqualsEmptyString()
    {
        bool result = this._evaluator.Evaluate(string.Empty, "a");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateCharacterExpressionEqualsNumeral()
    {
        bool result = this._evaluator.Evaluate(1d, "a");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateCharacterExpressionEqualsSameCharacter()
    {
        bool result = this._evaluator.Evaluate("a", "a");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateCharacterExpressionEqualsDifferentCharacter()
    {
        bool result = this._evaluator.Evaluate("b", "a");
        Assert.IsFalse(result);
    }

    #endregion

    #region CharacterWithOperator Expression Tests

    [TestMethod]
    public void EvaluateCharacterWithOperatorExpressionEqualNull()
    {
        bool result = this._evaluator.Evaluate(null, ">a");
        Assert.IsFalse(result);
        result = this._evaluator.Evaluate(null, "<a");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateCharacterWithOperatorExpressionEqualsEmptyString()
    {
        bool result = this._evaluator.Evaluate(string.Empty, ">a");
        Assert.IsFalse(result);
        result = this._evaluator.Evaluate(string.Empty, "<a");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateCharacterWithOperatorExpressionEqualsNumeral()
    {
        bool result = this._evaluator.Evaluate(1d, ">a");
        Assert.IsFalse(result);
        result = this._evaluator.Evaluate(1d, "<a");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateShouldHandleLeadingEqualOperatorAndWildCard()
    {
        bool result = this._evaluator.Evaluate("TEST", "=*EST*");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateCharacterWithOperatorExpressionEqualsSameCharacter()
    {
        bool result = this._evaluator.Evaluate("a", ">a");
        Assert.IsFalse(result);
        result = this._evaluator.Evaluate("a", ">=a");
        Assert.IsTrue(result);
        result = this._evaluator.Evaluate("a", "<a");
        Assert.IsFalse(result);
        result = this._evaluator.Evaluate("a", ">=a");
        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateCharacterWithOperatorExpressionEqualsDifferentCharacter()
    {
        bool result = this._evaluator.Evaluate("b", ">a");
        Assert.IsTrue(result);
        result = this._evaluator.Evaluate("b", "<a");
        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateCharacterWithSpaceBetweenOperatorAndCharacter()
    {
        bool result = this._evaluator.Evaluate("b", "> a");
        Assert.IsTrue(result);
        result = this._evaluator.Evaluate("b", "< a");
        Assert.IsFalse(result);
    }

    #endregion
}