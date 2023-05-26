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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class ParsingScopesTest
    {
        private ParsingScopes _parsingScopes;
        private IParsingLifetimeEventHandler _lifeTimeEventHandler;

        [TestInitialize]
        public void Setup()
        {
            this._lifeTimeEventHandler = A.Fake<IParsingLifetimeEventHandler>();
            this._parsingScopes = new ParsingScopes(this._lifeTimeEventHandler);
        }

        [TestMethod]
        public void CreatedScopeShouldBeCurrentScope()
        {
            using ParsingScope? scope = this._parsingScopes.NewScope(RangeAddress.Empty);
            Assert.AreEqual(this._parsingScopes.Current, scope);
        }

        [TestMethod]
        public void CurrentScopeShouldHandleNestedScopes()
        {
            using (ParsingScope? scope1 = this._parsingScopes.NewScope(RangeAddress.Empty))
            {
                Assert.AreEqual(this._parsingScopes.Current, scope1);
                using (ParsingScope? scope2 = this._parsingScopes.NewScope(RangeAddress.Empty))
                {
                    Assert.AreEqual(this._parsingScopes.Current, scope2);
                }
                Assert.AreEqual(this._parsingScopes.Current, scope1);
            }
            Assert.IsNull(this._parsingScopes.Current);
        }

        [TestMethod]
        public void CurrentScopeShouldBeNullWhenScopeHasTerminated()
        {
            using (ParsingScope? scope = this._parsingScopes.NewScope(RangeAddress.Empty))
            { }
            Assert.IsNull(this._parsingScopes.Current);
        }

        [TestMethod]
        public void NewScopeShouldSetParentOnCreatedScopeIfParentScopeExisted()
        {
            using ParsingScope? scope1 = this._parsingScopes.NewScope(RangeAddress.Empty);
            using ParsingScope? scope2 = this._parsingScopes.NewScope(RangeAddress.Empty);
            Assert.AreEqual(scope1, scope2.Parent);
        }

        [TestMethod]
        public void LifetimeEventHandlerShouldBeCalled()
        {
            using (ParsingScope? scope = this._parsingScopes.NewScope(RangeAddress.Empty))
            { }
            A.CallTo(() => this._lifeTimeEventHandler.ParsingCompleted()).MustHaveHappened();
        }
    }
}