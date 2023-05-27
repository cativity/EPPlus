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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database;

[TestClass]
public class RowMatcherTests
{
    private static ExcelDatabaseCriteria GetCriteria(Dictionary<ExcelDatabaseCriteriaField, object> items)
    {
        _ = A.Fake<ExcelDataProvider>();
        ExcelDatabaseCriteria? criteria = A.Fake<ExcelDatabaseCriteria>(); // (provider, string.Empty);

        _ = A.CallTo(() => criteria.Items).Returns(items);

        return criteria;
    }

    [TestMethod]
    public void IsMatchShouldReturnTrueIfCriteriasMatch()
    {
        ExcelDatabaseRow? data = new ExcelDatabaseRow();
        data["Crit1"] = 1;
        data["Crit2"] = 2;
        data["Crit3"] = 3;

        Dictionary<ExcelDatabaseCriteriaField, object>? crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
        crit[new ExcelDatabaseCriteriaField("Crit1")] = 1;
        crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

        RowMatcher? matcher = new RowMatcher();

        ExcelDatabaseCriteria? criteria = GetCriteria(crit);

        Assert.IsTrue(matcher.IsMatch(data, criteria));
    }

    [TestMethod]
    public void IsMatchShouldReturnFalseIfCriteriasDoesNotMatch()
    {
        ExcelDatabaseRow? data = new ExcelDatabaseRow();
        data["Crit1"] = 1;
        data["Crit2"] = 2;
        data["Crit3"] = 3;

        Dictionary<ExcelDatabaseCriteriaField, object>? crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
        crit[new ExcelDatabaseCriteriaField("Crit1")] = 1;
        crit[new ExcelDatabaseCriteriaField("Crit3")] = 4;

        RowMatcher? matcher = new RowMatcher();

        ExcelDatabaseCriteria? criteria = GetCriteria(crit);

        Assert.IsFalse(matcher.IsMatch(data, criteria));
    }

    [TestMethod]
    public void IsMatchShouldMatchStrings1()
    {
        ExcelDatabaseRow? data = new ExcelDatabaseRow();
        data["Crit1"] = "1";
        data["Crit2"] = 2;
        data["Crit3"] = 3;

        Dictionary<ExcelDatabaseCriteriaField, object>? crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
        crit[new ExcelDatabaseCriteriaField("Crit1")] = "1";
        crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

        RowMatcher? matcher = new RowMatcher();

        ExcelDatabaseCriteria? criteria = GetCriteria(crit);

        Assert.IsTrue(matcher.IsMatch(data, criteria));
    }

    [TestMethod]
    public void IsMatchShouldMatchStrings2()
    {
        ExcelDatabaseRow? data = new ExcelDatabaseRow();
        data["Crit1"] = "2";
        data["Crit2"] = 2;
        data["Crit3"] = 3;

        Dictionary<ExcelDatabaseCriteriaField, object>? crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
        crit[new ExcelDatabaseCriteriaField("Crit1")] = "1";
        crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

        RowMatcher? matcher = new RowMatcher();

        ExcelDatabaseCriteria? criteria = GetCriteria(crit);

        Assert.IsFalse(matcher.IsMatch(data, criteria));
    }

    [TestMethod]
    public void IsMatchShouldMatchWildcardStrings()
    {
        ExcelDatabaseRow? data = new ExcelDatabaseRow();
        data["Crit1"] = "test";
        data["Crit2"] = 2;
        data["Crit3"] = 3;

        Dictionary<ExcelDatabaseCriteriaField, object>? crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
        crit[new ExcelDatabaseCriteriaField("Crit1")] = "t*t";
        crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

        RowMatcher? matcher = new RowMatcher();

        ExcelDatabaseCriteria? criteria = GetCriteria(crit);

        Assert.IsTrue(matcher.IsMatch(data, criteria));
    }

    [TestMethod]
    public void IsMatchShouldMatchNumericExpression()
    {
        ExcelDatabaseRow? data = new ExcelDatabaseRow();
        data["Crit1"] = "test";
        data["Crit2"] = 2;
        data["Crit3"] = 3;

        Dictionary<ExcelDatabaseCriteriaField, object>? crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
        crit[new ExcelDatabaseCriteriaField("Crit2")] = "<3";
        crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

        RowMatcher? matcher = new RowMatcher();

        ExcelDatabaseCriteria? criteria = GetCriteria(crit);

        Assert.IsTrue(matcher.IsMatch(data, criteria));
    }

    [TestMethod]
    public void IsMatchShouldHandleFieldIndex()
    {
        ExcelDatabaseRow? data = new ExcelDatabaseRow();
        data["Crit1"] = "test";
        data["Crit2"] = 2;
        data["Crit3"] = 3;

        Dictionary<ExcelDatabaseCriteriaField, object>? crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
        crit[new ExcelDatabaseCriteriaField(2)] = "<3";
        crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

        RowMatcher? matcher = new RowMatcher();

        ExcelDatabaseCriteria? criteria = GetCriteria(crit);

        Assert.IsTrue(matcher.IsMatch(data, criteria));
    }
}