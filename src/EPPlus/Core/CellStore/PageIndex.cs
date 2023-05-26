/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Drawing;

namespace OfficeOpenXml.Core.CellStore;

internal class PageIndex : IndexBase, IDisposable
{
    public PageIndex(int pageSizeMin)
    {
        this.Rows = new IndexItem[pageSizeMin];
        this.RowCount = 0;
    }
    public PageIndex(IndexItem[] rows, int count)
    {
        this.Rows = rows;
        this.RowCount = count;
    }
    public PageIndex(PageIndex pageItem, int start, int size)
        : this(pageItem, start, size, pageItem.Index, pageItem.Offset)
    {

    }
    public PageIndex(PageIndex pageItem, int start, int size, short index, int offset, int arraySize = -1)
    {
        if(arraySize<0)
        {
            arraySize = CellStore<int>.GetSize(size);
        }

        this.Rows = new IndexItem[arraySize];
        Array.Copy(pageItem.Rows, start, this.Rows, 0, pageItem.RowCount-start);
        this.RowCount = size;
        this.Index = index;
        this.Offset = offset;
    }
    ~PageIndex()
    {
        this.Rows = null;
    }
    internal int Offset = 0;
    /// <summary>
    /// Rows in the rows collection. 
    /// </summary>
    internal int RowCount;
    internal int IndexOffset
    {
        get
        {
            return this.IndexExpanded + this.Offset;
        }
    }
    internal int IndexExpanded
    {
        get
        {
            return (this.Index << CellStoreSettings._pageBits);
        }
    }
    internal IndexItem[] Rows { get; set; }
    /// <summary>
    /// First row index minus last row index
    /// </summary>
    internal int RowSpan
    {
        get
        {
            return this.MaxIndex - this.MinIndex+1;
        }
    }

    internal int GetPosition(int offset)
    {
        return ArrayUtil.OptimizedBinarySearch(this.Rows, offset, this.RowCount);
    }
    internal int GetRowPosition(int row)
    {
        int offset = row - this.IndexOffset;
        return ArrayUtil.OptimizedBinarySearch(this.Rows, offset, this.RowCount);
    }
    internal int GetNextRow(int row)
    {
        int o = this.GetRowPosition(row);
        if (o < 0)
        {
            o = ~o;
            if (o < this.RowCount)
            {
                return o;
            }
            else
            {
                return -1;
            }
        }
        return o;
    }

    public int MinIndex
    {
        get
        {
            if (this.RowCount > 0)
            {
                return this.IndexOffset + this.Rows[0].Index;
            }
            else
            {
                return -1;
            }
        }
    }
    public int MaxIndex
    {
        get
        {
            if (this.RowCount > 0)
            {
                return this.IndexOffset + this.Rows[this.RowCount - 1].Index;
            }
            else
            {
                return -1;
            }
        }
    }
    public int GetIndex(int pos)
    {
        return this.IndexOffset + this.Rows[pos].Index;
    }
    public void Dispose()
    {
        this.Rows = null;
    }

    internal bool IsWithin(int fromRow, int toRow)
    {
        return fromRow <= this.MinIndex  && toRow >= this.MaxIndex;
    }
    internal bool StartsWithin(int fromRow, int toRow)
    {
        return fromRow <= this.MaxIndex && toRow >= this.MinIndex;
    }

    internal bool StartsAfter(int row)
    {
        return row > this.MaxIndex;
    }

    internal int GetRow(int rowIx)
    {
        return this.IndexOffset + this.Rows[rowIx].Index;
    }
}