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

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class IntegerExpressionTests
    {
        [TestMethod]
        public void MergeWithNextWithPlusOperatorShouldCalulateSumCorrectly()
        {
            IntegerExpression? exp1 = new IntegerExpression("1");
            exp1.Operator = Operator.Plus;
            IntegerExpression? exp2 = new IntegerExpression("2");
            exp1.Next = exp2;

            Expression? result = exp1.MergeWithNext();

            Assert.AreEqual(3d, result.Compile().Result);
        }

        [TestMethod]
        public void MergeWithNextWithPlusOperatorShouldSetNextPointer()
        {
            IntegerExpression? exp1 = new IntegerExpression("1");
            exp1.Operator = Operator.Plus;
            IntegerExpression? exp2 = new IntegerExpression("2");
            exp1.Next = exp2;

            Expression? result = exp1.MergeWithNext();

            Assert.IsNull(result.Next);
        }

        //[TestMethod]
        //public void CompileShouldHandlePercent()
        //{
        //    var exp1 = new IntegerExpression("1");
        //    exp1.Operator = Operator.Percent;
        //    exp1.Next = ConstantExpressions.Percent;
        //    var result = exp1.Compile();
        //    Assert.AreEqual(0.01, result.Result);
        //    Assert.AreEqual(DataType.Decimal, result.DataType);
        //}
    }
}
