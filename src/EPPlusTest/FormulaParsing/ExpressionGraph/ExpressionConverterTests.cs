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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System.Globalization;
using System.Threading;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExpressionConverterTests
    {
        private IExpressionConverter _converter;

        [TestInitialize]
        public void Setup()
        {
            this._converter = new ExpressionConverter();
        }

        [TestMethod]
        public void ToStringExpressionShouldConvertIntegerExpressionToStringExpression()
        {
            IntegerExpression? integerExpression = new IntegerExpression("2");
            StringExpression? result = this._converter.ToStringExpression(integerExpression);
            Assert.IsInstanceOfType(result, typeof(StringExpression));
            Assert.AreEqual("2", result.Compile().Result);
        }

        [TestMethod]
        public void ToStringExpressionShouldCopyOperatorToStringExpression()
        {
            IntegerExpression? integerExpression = new IntegerExpression("2");
            integerExpression.Operator = Operator.Plus;
            StringExpression? result = this._converter.ToStringExpression(integerExpression);
            Assert.AreEqual(integerExpression.Operator, result.Operator);
        }

        [TestMethod]
        public void ToStringExpressionShouldConvertDecimalExpressionToStringExpression()
        {
            DecimalExpression? decimalExpression = new DecimalExpression("2.5");
            StringExpression? result = this._converter.ToStringExpression(decimalExpression);
            Assert.IsInstanceOfType(result, typeof(StringExpression));
            Assert.AreEqual($"2{CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator}5", result.Compile().Result);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateIntegerExpressionIfCompileResultIsInteger()
        {
            CompileResult? compileResult = new CompileResult(1, DataType.Integer);
            Expression? result = this._converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(IntegerExpression));
            Assert.AreEqual(1d, result.Compile().Result);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateStringExpressionIfCompileResultIsString()
        {
            CompileResult? compileResult = new CompileResult("abc", DataType.String);
            Expression? result = this._converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(StringExpression));
            Assert.AreEqual("abc", result.Compile().Result);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateDecimalExpressionIfCompileResultIsDecimal()
        {
            CompileResult? compileResult = new CompileResult(2.5d, DataType.Decimal);
            Expression? result = this._converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(DecimalExpression));
            Assert.AreEqual(2.5d, result.Compile().Result);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateBooleanExpressionIfCompileResultIsBoolean()
        {
            CompileResult? compileResult = new CompileResult("true", DataType.Boolean);
            Expression? result = this._converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(BooleanExpression));
            Assert.IsTrue((bool)result.Compile().Result);
        }
    }
}
