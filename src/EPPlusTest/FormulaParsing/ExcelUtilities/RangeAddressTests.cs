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

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class RangeAddressTests
    {
        private RangeAddressFactory _factory;

        [TestInitialize]
        public void Setup()
        {
            ExcelDataProvider? provider = A.Fake<ExcelDataProvider>();
            _factory = new RangeAddressFactory(provider);
        }

        [TestMethod]
        public void CollideShouldReturnTrueIfRangesCollides()
        {
            RangeAddress? address1 = _factory.Create("A1:A6");
            RangeAddress? address2 = _factory.Create("A5");
            Assert.IsTrue(address1.CollidesWith(address2));
        }

        [TestMethod]
        public void CollideShouldReturnFalseIfRangesDoesNotCollide()
        {
            RangeAddress? address1 = _factory.Create("A1:A6");
            RangeAddress? address2 = _factory.Create("A8");
            Assert.IsFalse(address1.CollidesWith(address2));
        }

        [TestMethod]
        public void CollideShouldReturnFalseIfRangesCollidesButWorksheetNameDiffers()
        {
            RangeAddress? address1 = _factory.Create("Ws!A1:A6");
            RangeAddress? address2 = _factory.Create("A5");
            Assert.IsFalse(address1.CollidesWith(address2));
        }
    }
}
