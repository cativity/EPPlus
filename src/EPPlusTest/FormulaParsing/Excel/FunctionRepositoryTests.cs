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
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace EPPlusTest.FormulaParsing.Excel.Functions;

[TestClass]
public class FunctionRepositoryTests
{
    #region LoadModule Tests

    [TestMethod]
    public void LoadModulePopulatesFunctionsAndCustomCompilers()
    {
        FunctionRepository? functionRepository = FunctionRepository.Create();
        Assert.IsFalse(functionRepository.IsFunctionName(MyFunction.Name));
        Assert.IsFalse(functionRepository.CustomCompilers.ContainsKey(typeof(MyFunction)));
        functionRepository.LoadModule(new TestFunctionModule());
        Assert.IsTrue(functionRepository.IsFunctionName(MyFunction.Name));
        Assert.IsTrue(functionRepository.CustomCompilers.ContainsKey(typeof(MyFunction)));

        // Make sure reloading the module overwrites previous functions and compilers
        functionRepository.LoadModule(new TestFunctionModule());
    }

    #endregion

    #region Nested Classes

    public class TestFunctionModule : FunctionsModule
    {
        public TestFunctionModule()
        {
            MyFunction? myFunction = new MyFunction();
            MyFunctionCompiler? customCompiler = new MyFunctionCompiler(myFunction, ParsingContext.Create());
            this.Functions.Add(MyFunction.Name, myFunction);
            this.CustomCompilers.Add(typeof(MyFunction), customCompiler);
        }
    }

    public class MyFunction : ExcelFunction
    {
        public const string Name = "MyFunction";

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context) => throw new NotImplementedException();
    }

    public class MyFunctionCompiler : FunctionCompiler
    {
        public MyFunctionCompiler(MyFunction function, ParsingContext context)
            : base(function, context)
        {
        }

        public override CompileResult Compile(IEnumerable<Expression> children) => throw new NotImplementedException();
    }

    #endregion
}