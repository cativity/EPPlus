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

using OfficeOpenXml.FormulaParsing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusTest.FormulaParsing;

[TestClass]
public class ExcelAddressCacheTests
{
    [TestMethod]
    public void ShouldGenerateNewIds()
    {
        ExcelAddressCache? cache = new ExcelAddressCache();
        int firstId = cache.GetNewId();
        Assert.AreEqual(1, firstId);

        int secondId = cache.GetNewId();
        Assert.AreEqual(2, secondId);
    }

    [TestMethod]
    public void ShouldReturnCachedAddress()
    {
        ExcelAddressCache? cache = new ExcelAddressCache();
        int id = cache.GetNewId();
        string? address = "A1";
        bool result = cache.Add(id, address);
        Assert.IsTrue(result);
        Assert.AreEqual(address, cache.Get(id));
    }

    [TestMethod]
    public void AddShouldReturnFalseIfUsedId()
    {
        ExcelAddressCache? cache = new ExcelAddressCache();
        int id = cache.GetNewId();
        string? address = "A1";
        bool result = cache.Add(id, address);
        Assert.IsTrue(result);
        bool result2 = cache.Add(id, address);
        Assert.IsFalse(result2);
    }

    [TestMethod]
    public void ClearShouldResetId()
    {
        ExcelAddressCache? cache = new ExcelAddressCache();
        int id = cache.GetNewId();
        Assert.AreEqual(1, id);
        string? address = "A1";
        _ = cache.Add(id, address);
        Assert.AreEqual(1, cache.Count);
        int id2 = cache.GetNewId();
        Assert.AreEqual(2, id2);
        cache.Clear();
        int id3 = cache.GetNewId();
        Assert.AreEqual(1, id3);
    }
}