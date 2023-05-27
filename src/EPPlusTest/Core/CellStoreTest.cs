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
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Style;

namespace EPPlusTest.Core;

[TestClass]
public class CellStoreTest : TestBase
{
    //public const int _cellBits = 6;
    [ClassInitialize]
    public static void Init(TestContext context)
    {
        //CellStoreSettings.InitSize(_cellBits);
    }

    #region SetValue

    [TestMethod]
    public void AddRandomRows()
    {
        CellStore<object>? cellStore = new CellStore<object>();
        cellStore.SetValue(25000, 1, 25000);
        cellStore.SetValue(1200, 1, 1200);
        cellStore.SetValue(1025000, 1, 1025000);

        Assert.AreEqual(25000, cellStore.GetValue(25000, 1));
        Assert.AreEqual(1200, cellStore.GetValue(1200, 1));
        Assert.AreEqual(1025000, cellStore.GetValue(1025000, 1));
    }

    [TestMethod]
    public void ForParallelSet()
    {
        List<int>? lst = new List<int>();
        CellStore<object>? cellStore = new CellStore<object>();

        for (int i = 0; i < 100000; i++)
        {
            lst.Add(i);
        }

        ParallelLoopResult r = Parallel.ForEach(lst,
                                                l =>
                                                {
                                                    cellStore.SetValue(l + 1, 1, l + 1);
                                                    cellStore.SetValue(l + 1, 2, $"Value {l + 1}");
                                                });

        while (r.IsCompleted == false)
        {
            Thread.Sleep(1000);
        }

        ;

        Assert.AreEqual(15, cellStore.GetValue(15, 1));
        Assert.AreEqual("Value 15", cellStore.GetValue(15, 2));
        Assert.AreEqual(9999, cellStore.GetValue(9999, 1));
        Assert.AreEqual("Value 9999", cellStore.GetValue(9999, 2));
        Assert.AreEqual(99999, cellStore.GetValue(99999, 1));
        Assert.AreEqual("Value 99999", cellStore.GetValue(99999, 2));
    }

    [TestMethod]
    public void ForParallelDelete()
    {
        List<int>? lst = new List<int>();
        CellStore<object>? cellStore = new CellStore<object>();
        int maxRow = 100000;

        for (int i = 0; i < maxRow; i++)
        {
            cellStore.SetValue(i, 0, i + 1);
            cellStore.SetValue(i, 2, $"Value {i + 1}");
        }

        ParallelLoopResult r = Parallel.For(0, maxRow, l => { cellStore.Delete(l, 0, 1, 0); });
    }

    #endregion

    #region Delete

    [TestMethod]
    public void DeletePrevRowWhenCreatePage()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        //Insert second page first default row, when cellBits is 5.
        cellStore.SetValue(1, 1, 1);
        cellStore.SetValue(33, 1, 33);

        //Delete prev row, and shift back
        cellStore.Delete(30, 1, 2, ExcelPackage.MaxColumns);

        Assert.AreEqual(33, cellStore.GetValue(31, 1));
    }

    [TestMethod]
    public void DeleteFromStartPageThreeRowsEveryRow()
    {
        CellStore<int>? cellStore = new CellStore<int>();
        LoadCellStore(cellStore, 100, 1500);

        for (int i = 1; i < 500; i++)
        {
            cellStore.Delete(1, 1, 3, ExcelPackage.MaxColumns); //Delete three rows each time.
            int row100 = 100 - (i * 3);

            if (row100 > 0)
            {
                Assert.AreEqual(100, cellStore.GetValue(row100, 1));
            }
            else
            {
                Assert.AreEqual(100 - row100 + 1, cellStore.GetValue(1, 1));
            }
        }
    }

    [TestMethod]
    public void DeleteFromStartPageThreeRowsEveryOtherRow()
    {
        CellStore<int>? cellStore = new CellStore<int>();
        LoadCellStore(cellStore, 100, 2200, 2);

        for (int i = 1; i < 500; i += 1)
        {
            cellStore.Delete(1, 1, 3, ExcelPackage.MaxColumns); //Delete three rows each time.
            int row100 = 100 - (i * 3);

            if (row100 > 0)
            {
                Assert.AreEqual(100, cellStore.GetValue(row100, 1));
            }
            else
            {
                int r = 100 - row100 + 1;

                if (r % 2 == 0)
                {
                    Assert.AreEqual(r, cellStore.GetValue(1, 1));
                }
                else
                {
                    Assert.AreEqual(0, cellStore.GetValue(1, 1));
                }
            }
        }
    }

    [TestMethod]
    public void DeleteFromStartPageThreeRowsEveryThirdRow()
    {
        CellStore<int>? cellStore = new CellStore<int>();
        LoadCellStore(cellStore, 100, 3300, 3);

        for (int i = 1; i < 500; i += 1)
        {
            cellStore.Delete(1, 1, 3, ExcelPackage.MaxColumns); //Delete three rows each time.
            int row100 = 100 - (i * 3);

            if (row100 > 0)
            {
                Assert.AreEqual(100, cellStore.GetValue(row100, 1));
            }
            else
            {
                int r = 100 - row100 + 1;

                if ((r % 3) - 1 == 0)
                {
                    Assert.AreEqual(r, cellStore.GetValue(1, 1));
                }
                else
                {
                    Assert.AreEqual(0, cellStore.GetValue(1, 1));
                }
            }
        }
    }

    [TestMethod]
    public void DeleteFromStartPageThreeRowsEveryRowWithRowOneSet()
    {
        CellStore<int>? cellStore = new CellStore<int>();
        cellStore.SetValue(1, 1, 1);
        LoadCellStore(cellStore, 100, 1500);

        for (int i = 1; i < 500; i++)
        {
            cellStore.Delete(2, 1, 3, ExcelPackage.MaxColumns); //Delete three rows each time.
            int row100 = 100 - (i * 3);

            if (row100 > 1)
            {
                Assert.AreEqual(100, cellStore.GetValue(row100, 1));
            }
            else
            {
                Assert.AreEqual(100 - row100 + 2, cellStore.GetValue(2, 1));
            }
        }
    }

    [TestMethod]
    public void DeleteMerge()
    {
        //Setup
        CellStore<int>? cellStore = new CellStore<int>();
        int pageSize = 1 << CellStoreSettings._pageBits;
        cellStore.SetValue(2, 1, 2); //Set Row 2;

        int row1 = pageSize + 1;
        cellStore.SetValue(row1, 1, row1);
        int row2 = (pageSize * 2) + 1;
        cellStore.SetValue(row2, 1, row2);

        //Act
        cellStore.Delete(3, 1, pageSize - 2, ExcelPackage.MaxColumns);
        cellStore.Delete(4, 1, pageSize - 1, ExcelPackage.MaxColumns);

        //Assert
        Assert.AreEqual(1, cellStore.ColumnCount);
        Assert.AreEqual(2, cellStore._columnIndex[0].PageCount);
        Assert.AreEqual(2, cellStore.GetValue(2, 1));
        Assert.AreEqual(row1, cellStore.GetValue(3, 1));
        Assert.AreEqual(row2, cellStore.GetValue(4, 1));
    }

    [TestMethod]
    public void DeleteRow2_3()
    {
        //Setup
        CellStore<int>? cellStore = new CellStore<int>();
        cellStore.SetValue(3, 1, 1);
        cellStore.SetValue(4, 1, 2);
        cellStore.SetValue(5, 1, 3);
        cellStore.SetValue(6, 1, 4);

        cellStore.Delete(2, 1, 2, 1);

        Assert.AreEqual(2, cellStore.GetValue(2, 1));
    }

    #endregion

    #region Clear

    [TestMethod]
    public void ClearInsideAndOverPage()
    {
        //Setup
        CellStore<int>? cellStore = new CellStore<int>();
        LoadCellStore(cellStore, 1, 300);

        cellStore.Clear(2, 1, 3, ExcelPackage.MaxColumns);

        //Clear from 2-4
        Assert.AreEqual(1, cellStore.GetValue(1, 1));
        Assert.AreEqual(0, cellStore.GetValue(2, 1));
        Assert.AreEqual(0, cellStore.GetValue(3, 1));
        Assert.AreEqual(0, cellStore.GetValue(4, 1));
        Assert.AreEqual(5, cellStore.GetValue(5, 1));

        //Clear from 3-7
        cellStore.Clear(3, 1, 5, ExcelPackage.MaxColumns);
        Assert.AreEqual(0, cellStore.GetValue(5, 1));
        Assert.AreEqual(0, cellStore.GetValue(7, 1));
        Assert.AreEqual(8, cellStore.GetValue(8, 1));

        //Clear from 10-44
        cellStore.Clear(10, 1, 35, ExcelPackage.MaxColumns);
        Assert.AreEqual(9, cellStore.GetValue(9, 1));
        Assert.AreEqual(0, cellStore.GetValue(10, 1));
        Assert.AreEqual(0, cellStore.GetValue(44, 1));
        Assert.AreEqual(45, cellStore.GetValue(45, 1));

        //Clear from 50-211
        cellStore.Clear(50, 1, 162, ExcelPackage.MaxColumns);
        Assert.AreEqual(49, cellStore.GetValue(49, 1));
        Assert.AreEqual(0, cellStore.GetValue(50, 1));
        Assert.AreEqual(0, cellStore.GetValue(211, 1));
        Assert.AreEqual(212, cellStore.GetValue(212, 1));
        Assert.AreEqual(250, cellStore.GetValue(250, 1));
    }

    #endregion

    #region Insert

    [TestMethod]
    public void InsertAndDeleteRowsOnPage5Bits()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        LoadCellStore(cellStore, 1, 10000, 1, 3);
        Assert.AreEqual(5000, cellStore.GetValue(5000, 1));
        Assert.AreEqual(10000, cellStore.GetValue(5000, 2));
        Assert.AreEqual(15000, cellStore.GetValue(5000, 3));

        //Insert 32 rows
        int InsertFrom1 = 32;
        int insertRows1 = 64;
        cellStore.Insert(InsertFrom1, 1, insertRows1, 0);

        Assert.AreEqual(InsertFrom1 - 1, cellStore.GetValue(InsertFrom1 - 1, 1));
        Assert.AreEqual(default(int), cellStore.GetValue(InsertFrom1, 1));
        Assert.AreEqual(default(int), cellStore.GetValue(InsertFrom1 + insertRows1 - 1, 1));
        Assert.AreEqual(32, cellStore.GetValue(InsertFrom1 + insertRows1, 1));

        cellStore.SetValue(32, 1, 10032);
        cellStore.SetValue(33, 1, 10033);
        cellStore.SetValue(34, 1, 10033);
    }

    [TestMethod]
    public void InsertRowEveryOther()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        LoadCellStore(cellStore, 1, 1000, 1);

        for (int r = 1; r < 1000; r++)
        {
            int row = ((r - 1) * 2) + 1;
            cellStore.Insert(row + 1, 1, 1, 0);
            Assert.AreEqual(r, cellStore.GetValue(row, 1));
        }
    }

    [TestMethod]
    public void InsertTwoRowsEveryThird()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        LoadCellStore(cellStore, 1, 1000, 1);

        for (int r = 1; r < 1000; r++)
        {
            int row = ((r - 1) * 3) + 1;
            cellStore.Insert(row + 1, 1, 2, 0);
            Assert.AreEqual(r, cellStore.GetValue(row, 1));
        }
    }

    [TestMethod]
    public void InsertThreeRowsEveryForth()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        LoadCellStore(cellStore, 1, 5000, 1);

        for (int r = 1; r < 5000; r++)
        {
            int row = ((r - 1) * 4) + 1;
            cellStore.Insert(row + 1, 1, 3, 0);
            Assert.AreEqual(r, cellStore.GetValue(row, 1));
        }
    }

    [TestMethod]
    public void InsertFourRowsEveryFifth()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        LoadCellStore(cellStore, 1, 5000, 1);

        for (int r = 1; r < 5000; r++)
        {
            int row = ((r - 1) * 5) + 1;
            cellStore.Insert(row + 1, 1, 4, 0);
            Assert.AreEqual(r, cellStore.GetValue(row, 1));
        }
    }

    [TestMethod]
    public void Insert1To500RowsFromStart()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        LoadCellStore(cellStore, 1, 5000, 1);

        int row = 1;

        for (int r = 1; r < 500; r++)
        {
            Assert.AreEqual(r, cellStore.GetValue(row, 1));
            cellStore.Insert(row, 1, r, 0);
            row += r + 1;
        }
    }

    [TestMethod]
    public void Delete1To500RowsFromStart()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        LoadCellStore(cellStore, 1, 50000, 1);

        int v = 1;

        for (int r = 1; r < 300; r++)
        {
            Assert.AreEqual(v, cellStore.GetValue(v, 1));
            cellStore.Delete(r, 1, r, 0);
            v += r + 1;
        }
    }

    [TestMethod]
    public void Delete1To500RowsFromStartThenClear()
    {
        CellStore<int>? cellStore = new CellStore<int>();

        LoadCellStore(cellStore, 1, 50000, 1);

        int v = 1;

        for (int r = 1; r < 300; r++)
        {
            Assert.AreEqual(v, cellStore.GetValue(v, 1));
            cellStore.Delete(r, 1, r, 0);
            v += r + 1;
        }
    }

    [TestMethod]
    public void ValidatePerformance()
    {
        CellStore<int>? cellStore = new CellStore<int>();
        TimeSpan acceptable = new TimeSpan(0, 0, 1);
        DateTime dt = DateTime.Now;
        LoadCellStore(cellStore, 1, 1000000);
        TimeSpan elapsedTime = DateTime.Now - dt;
        Assert.IsTrue(elapsedTime < acceptable, "Cellstore performance is slow");
    }

    [TestMethod]
    public void Add35000RowsAtOnce()
    {
        CellStore<int>? cellStore = new CellStore<int>();
        cellStore.SetValue(1, 1, 1);
        cellStore.SetValue(2, 1, 2);
        cellStore.SetValue(10000, 1, 10000);

        cellStore.Insert(2, 1, 35000, 1);

        Assert.AreEqual(1, cellStore.GetValue(1, 1));
        Assert.AreEqual(2, cellStore.GetValue(35002, 1));
        Assert.AreEqual(10000, cellStore.GetValue(45000, 1));
    }

    #endregion

    private static void LoadCellStore(CellStore<int> cellStore, int fromRow = 1, int toRow = 1000, int add = 1, int cols = 1)
    {
        for (int row = fromRow; row <= toRow; row += add)
        {
            for (int col = 1; col <= cols; col++)
            {
                cellStore.SetValue(row, col, row * col);
            }
        }
    }
}