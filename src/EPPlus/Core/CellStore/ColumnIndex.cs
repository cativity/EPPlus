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
using System.Collections.Generic;

namespace OfficeOpenXml.Core.CellStore
{
    internal class ColumnIndex<T> : IndexBase, IDisposable
    {
        internal List<T> _values = new List<T>();

        public ColumnIndex()
        {
            this._pages = new PageIndex[CellStoreSettings.PagesPerColumnMin];
            this.PageCount = 0;
        }
        ~ColumnIndex()
        {
            this._pages = null;
        }
        internal int GetPagePosition(int Row)
        {
            int page = (Row >> CellStoreSettings._pageBits);
            int pagePos;
            if (page >= 0 && page < this.PageCount && this._pages[page].Index == page)
            {
                pagePos = page;
            }
            else
            {
                pagePos = ArrayUtil.OptimizedBinarySearch(this._pages, page, this.PageCount);
            }

            if (pagePos >= 0)
            {
                this.GetPage(Row, ref pagePos);
                return pagePos;
            }
            else
            {
                int p = ~pagePos;

                if (this.GetPage(Row, ref p))
                {
                    return p;
                }
                else
                {
                    return pagePos;
                }
            }
        }

        private bool GetPage(int Row, ref int pagePos)
        {
            if (pagePos < this.PageCount && this._pages[pagePos].MinIndex <= Row && (pagePos + 1 == this.PageCount || this._pages[pagePos + 1].MinIndex > Row))
            {
                return true;
            }
            else
            {
                if (pagePos + 1 < this.PageCount && (this._pages[pagePos + 1].MinIndex <= Row))
                {
                    do
                    {
                        pagePos++;
                    }
                    while (pagePos + 1 < this.PageCount && this._pages[pagePos + 1].MinIndex <= Row);
                    return true;
                }
                else if (pagePos - 1 >= 0 && this._pages[pagePos - 1].MaxIndex >= Row)
                {
                    do
                    {
                        pagePos--;
                    }
                    while (pagePos - 1 > 0 && this._pages[pagePos - 1].MaxIndex >= Row);
                    return true;
                }
                return false;
            }
        }
        internal int GetNextRow(int row)
        {
            int p = this.GetPagePosition(row);
            if (p < 0)
            {
                p = ~p;
                if (p >= this.PageCount)
                {
                    return -1;
                }
                else
                {

                    if (this._pages[p].IndexOffset + this._pages[p].Rows[0].Index < row)
                    {
                        if (p + 1 >= this.PageCount)
                        {
                            return -1;
                        }
                        else
                        {
                            return this._pages[p + 1].IndexOffset + this._pages[p].Rows[0].Index;
                        }
                    }
                    else
                    {
                        return this._pages[p].IndexOffset + this._pages[p].Rows[0].Index;
                    }
                }
            }
            else
            {
                if (p < this.PageCount)
                {
                    int r = this._pages[p].GetNextRow(row);
                    if (r >= 0)
                    {
                        return this._pages[p].IndexOffset + this._pages[p].Rows[r].Index;
                    }
                    else
                    {
                        if (++p < this.PageCount)
                        {
                            return this._pages[p].IndexOffset + this._pages[p].Rows[0].Index;
                        }
                        else
                        {
                            return -1;
                        }
                    }
                }
                else
                {
                    return -1;
                }
            }
        }
        internal int GetPointer(int Row)
        {
            int pos = this.GetPagePosition(Row);
            if (pos >= 0 && pos < this.PageCount)
            {
                PageIndex? pageItem = this._pages[pos];
                if (pageItem.MinIndex > Row)
                {
                    pos--;
                    if (pos < 0)
                    {
                        return -1;
                    }
                    else
                    {
                        pageItem = this._pages[pos];
                    }
                }
                int ix = Row - pageItem.IndexOffset;
                int cellPos = ArrayUtil.OptimizedBinarySearch(pageItem.Rows, ix, pageItem.RowCount);
                if (cellPos >= 0)
                {
                    return pageItem.Rows[cellPos].IndexPointer;
                }
                else //Cell does not exist
                {
                    return -1;
                }
            }
            else //Page does not exist
            {
                return -1;
            }
        }

        //internal int FindNext(int Page)
        //{
        //    var p = GetPagePosition(Page);
        //    if (p < 0)
        //    {
        //        return ~p;
        //    }
        //    return p;
        //}
        internal PageIndex[] _pages;
        internal int PageCount;
        public void Dispose()
        {
            if (this._pages == null)
            {
                return;
            }

            for (int p = 0; p < this.PageCount; p++)
            {
                (this._pages[p] as IDisposable)?.Dispose();
            }

            this._pages = null;
            if (this._values != null)
            {
                this._values.Clear();
            }
        }
    }
}