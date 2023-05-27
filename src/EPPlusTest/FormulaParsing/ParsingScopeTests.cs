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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing;

[TestClass]
public class ParsingScopeTests
{
    private IParsingLifetimeEventHandler _lifeTimeEventHandler;
    private ParsingScopes _parsingScopes;
    private RangeAddressFactory _factory;

    [TestInitialize]
    public void Setup()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        this._factory = new RangeAddressFactory(provider);
        this._lifeTimeEventHandler = A.Fake<IParsingLifetimeEventHandler>();
        this._parsingScopes = A.Fake<ParsingScopes>();
    }

    [TestMethod]
    public void ConstructorShouldSetAddress()
    {
        RangeAddress? expectedAddress = this._factory.Create("A1");
        ParsingScope? scope = new ParsingScope(this._parsingScopes, expectedAddress);
        Assert.AreEqual(expectedAddress, scope.Address);
    }

    [TestMethod]
    public void ConstructorShouldSetParent()
    {
        ParsingScope? parent = new ParsingScope(this._parsingScopes, this._factory.Create("A1"));
        ParsingScope? scope = new ParsingScope(this._parsingScopes, parent, this._factory.Create("A2"));
        Assert.AreEqual(parent, scope.Parent);
    }

    [TestMethod]
    public void ScopeShouldCallKillScopeOnDispose()
    {
        ParsingScope? scope = new ParsingScope(this._parsingScopes, this._factory.Create("A1"));
        ((IDisposable)scope).Dispose();
        _ = A.CallTo(() => this._parsingScopes.KillScope(scope)).MustHaveHappened();
    }
}