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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.ExcelUtilities;

[TestClass]
public class RangeAddressFactoryTests
{
    private RangeAddressFactory _factory;
    private const int ExcelMaxRows = 1048576;

    [TestInitialize]
    public void Setup()
    {
        ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
        _ = A.CallTo(() => provider.ExcelMaxRows).Returns(ExcelMaxRows);
        this._factory = new RangeAddressFactory(provider);
    }

    [TestMethod, ExpectedException(typeof(ArgumentException))]
    public void CreateShouldThrowIfSuppliedAddressIsNull() => _ = this._factory.Create(null);

    [TestMethod]
    public void CreateShouldReturnAndInstanceWithColPropertiesSet()
    {
        RangeAddress? address = this._factory.Create("A2");
        Assert.AreEqual(1, address.FromCol, "FromCol was not 1");
        Assert.AreEqual(1, address.ToCol, "ToCol was not 1");
    }

    [TestMethod]
    public void CreateShouldReturnAndInstanceWithRowPropertiesSet()
    {
        RangeAddress? address = this._factory.Create("A2");
        Assert.AreEqual(2, address.FromRow, "FromRow was not 2");
        Assert.AreEqual(2, address.ToRow, "ToRow was not 2");
    }

    [TestMethod]
    public void CreateShouldReturnAnInstanceWithFromAndToColSetWhenARangeAddressIsSupplied()
    {
        RangeAddress? address = this._factory.Create("A1:B2");
        Assert.AreEqual(1, address.FromCol);
        Assert.AreEqual(2, address.ToCol);
    }

    [TestMethod]
    public void CreateShouldReturnAnInstanceWithFromAndToRowSetWhenARangeAddressIsSupplied()
    {
        RangeAddress? address = this._factory.Create("A1:B3");
        Assert.AreEqual(1, address.FromRow);
        Assert.AreEqual(3, address.ToRow);
    }

    [TestMethod]
    public void CreateShouldSetWorksheetNameIfSuppliedInAddress()
    {
        RangeAddress? address = this._factory.Create("Ws!A1");
        Assert.AreEqual("Ws", address.Worksheet);
    }

    [TestMethod]
    public void CreateShouldReturnAnInstanceWithStringAddressSet()
    {
        RangeAddress? address = this._factory.Create(1, 1);
        Assert.AreEqual("A1", address.ToString());
    }

    [TestMethod]
    public void CreateShouldReturnAnInstanceWithFromAndToColSet()
    {
        RangeAddress? address = this._factory.Create(1, 0);
        Assert.AreEqual(1, address.FromCol);
        Assert.AreEqual(1, address.ToCol);
    }

    [TestMethod]
    public void CreateShouldReturnAnInstanceWithFromAndToRowSet()
    {
        RangeAddress? address = this._factory.Create(0, 1);
        Assert.AreEqual(1, address.FromRow);
        Assert.AreEqual(1, address.ToRow);
    }

    [TestMethod]
    public void CreateShouldReturnAnInstanceWithWorksheetSetToEmptyString()
    {
        RangeAddress? address = this._factory.Create(0, 1);
        Assert.AreEqual(string.Empty, address.Worksheet);
    }

    [TestMethod]
    public void CreateShouldReturnEntireColumnRangeWhenNoRowsAreSpecified()
    {
        RangeAddress? address = this._factory.Create("A:B");
        Assert.AreEqual(1, address.FromCol);
        Assert.AreEqual(2, address.ToCol);
        Assert.AreEqual(1, address.FromRow);
        Assert.AreEqual(ExcelMaxRows, address.ToRow);
    }
}