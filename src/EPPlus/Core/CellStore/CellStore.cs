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
namespace OfficeOpenXml.Core.CellStore
{
    /// <summary>
    /// For testing purpose only. Can be removed when cells store is fully optimized.
    /// </summary>
    internal static class CellStoreSettings
    {
        /**** Size constants ****/
        internal const int _pageBits = 13;   // 13bits = 8192  Note: Maximum is 13 bits since short is used (PageMax=16K)
        internal const int _pageSize = 1 << _pageBits;
        internal const int _pageSizeMax = _pageSize << 1;
        internal const int _pageSizeMin = 64; //1 << _pageBits;
        internal const int ColSizeMin = 32;
        internal const int PagesPerColumnMin = 32;

        //internal static void InitSize(int size)
        //{
        //    _pageBits = size;   // 13bits = 8192  Note: Maximum is 13 bits since short is used (PageMax=16K)
        //    _pageSize = 1 << _pageBits;
        //    _pageSizeMax = _pageSize << 1;
        //    _pageSizeMin = 1 << _pageBits;
        //    ColSizeMin = 32;
        //    PagesPerColumnMin = 32;
        //}
        //internal static void DefaultSize(int size)
        //{
        //    InitSize(13);
        //}
    }

    /// <summary>
    /// This is the store for all Rows, Columns and Cells.
    /// It is a Dictionary implementation that allows you to change the Key.
    /// Rows and Column data is stored in column with index 0(row data) and row with index 0 (column data).
    /// </summary>
    internal class CellStore<T> : IDisposable
    {
        internal ColumnIndex<T>[] _columnIndex;
        internal int ColumnCount;
        /// <summary>
        /// For internal use only. 
        /// Must be set before any instance of the CellStore is created.
        /// </summary>
        public CellStore()
        {
            _columnIndex = new ColumnIndex<T>[CellStoreSettings.ColSizeMin];
        }
        ~CellStore()
        {
            _columnIndex = null;
        }
        internal bool HasValues
        {
            get
            {
                foreach (ColumnIndex<T>? col in _columnIndex)
                {
                    if (col == null)
                    {
                        break;
                    }

                    if (col._values.Count > 0)
                    {
                        return true;
                    }
                }
                return false;
            }
        }
        internal int GetClosestColumnPosition(int column)
        {
            int pos = GetColumnPosition(column);
            if (pos < 0)
            {
                return ~pos;
            }
            return pos;
        }
        internal int GetColumnPosition(int column)
        {
            return ArrayUtil.OptimizedBinarySearch(_columnIndex, column, ColumnCount);
        }
        internal ColumnIndex<T> GetColumnIndex(int column)
        {
            int pos = ArrayUtil.OptimizedBinarySearch(_columnIndex, column, ColumnCount);
            if (pos >= 0 && pos <= ColumnCount)
            {
                return _columnIndex[pos];
            }
            return null;
        }
        internal CellStore<T> Clone()
        {
            int row, col;
            CellStore<T>? ret = new CellStore<T>();
            for (int c = 0; c < ColumnCount; c++)
            {
                ColumnIndex<T>? colIx = _columnIndex[c];
                col = colIx.Index;
                for (int p = 0; p < colIx.PageCount; p++)
                {
                    for (int r = 0; r < colIx._pages[p].RowCount; r++)
                    {
                        row = colIx._pages[p].IndexOffset + colIx._pages[p].Rows[r].Index;
                        ret.SetValue(row, col, colIx._values[colIx._pages[p].Rows[r].IndexPointer]);
                    }
                }
            }
            return ret;
        }
        internal int Count
        {
            get
            {
                int count = 0;
                for (int c = 0; c < ColumnCount; c++)
                {
                    for (int p = 0; p < _columnIndex[c].PageCount; p++)
                    {
                        count += _columnIndex[c]._pages[p].RowCount;
                    }
                }
                return count;
            }
        }

        //internal int Capacity
        //{
        //    get => _values.Capacity;
        //    set => _values.Capacity = value;
        //}

        internal bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
        {
            if (ColumnCount == 0)
            {
                fromRow = fromCol = toRow = toCol = 0;
                return false;
            }
            else
            {
                fromCol = _columnIndex[0].Index;
                int fromIndex = 0;
                if (fromCol <= 0 && ColumnCount > 1)
                {
                    fromCol = _columnIndex[1].Index;
                    fromIndex = 1;
                }
                else if (ColumnCount == 1 && fromCol <= 0)
                {
                    fromRow = fromCol = toRow = toCol = 0;
                    return false;
                }
                int col = ColumnCount - 1;
                while (col > 0)
                {
                    if (_columnIndex[col].PageCount == 0 || _columnIndex[col]._pages[0].RowCount > 1 || _columnIndex[col]._pages[0].Rows[0].Index > 0)
                    {
                        break;
                    }
                    col--;
                }
                toCol = _columnIndex[col].Index;
                if (toCol == 0)
                {
                    fromRow = fromCol = toRow = toCol = 0;
                    return false;
                }
                fromRow = toRow = 0;

                for (int c = fromIndex; c < ColumnCount; c++)
                {
                    int first, last;
                    if (_columnIndex[c].PageCount == 0)
                    {
                        continue;
                    }

                    if (_columnIndex[c]._pages[0].RowCount > 0 && _columnIndex[c]._pages[0].Rows[0].Index >= 0 &&
                        _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[0].Rows[0].Index > 0)
                    {
                        first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[0].Rows[0].Index;
                    }
                    else
                    {
                        if (_columnIndex[c]._pages[0].RowCount > 1)
                        {
                            first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[0].Rows[1].Index;
                        }
                        else if (_columnIndex[c].PageCount > 1)
                        {
                            first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[1].Rows[0].Index;
                        }
                        else
                        {
                            first = 0;
                        }
                    }
                    int lp = _columnIndex[c].PageCount - 1;
                    while (_columnIndex[c]._pages[lp].RowCount == 0 && lp != 0)
                    {
                        lp--;
                    }
                    PageIndex? p = _columnIndex[c]._pages[lp];
                    if (p.RowCount > 0)
                    {
                        last = p.IndexOffset + p.Rows[p.RowCount - 1].Index;
                    }
                    else
                    {
                        last = first;
                    }
                    if (first > 0 && (first < fromRow || fromRow == 0))
                    {
                        fromRow = first;
                    }
                    if (first > 0 && (last > toRow || toRow == 0))
                    {
                        toRow = last;
                    }
                }
                if (fromRow <= 0 || toRow <= 0)
                {
                    fromRow = fromCol = toRow = toCol = 0;
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
        internal int FindNext(int Column)
        {
            int c = GetColumnPosition(Column);
            if (c < 0)
            {
                return ~c;
            }
            return c;
        }
        internal T GetValue(int Row, int Column)
        {
            ColumnIndex<T>? c = GetColumnIndex(Column);
            if (c != null)
            {
                int i = c.GetPointer(Row);
                if (i >= 0)
                {
                    return c._values[i];
                }
            }
            return default(T);
        }
        internal bool Exists(int Row, int Column)
        {
            ColumnIndex<T>? c = GetColumnIndex(Column);
            if (c == null)
            {
                return false;
            }

            return c.GetPointer(Row) >= 0;
        }

        internal bool Exists(int Row, int Column, ref T value)
        {
            ColumnIndex<T>? c = GetColumnIndex(Column);
            if (c == null)
            {
                return false;
            }

            int p = c.GetPointer(Row);
            if (p >= 0)
            {
                value = c._values[p];
                return true;
            }
            else
            {
                return false;
            }
        }
        internal void SetValue(int row, int column, T value)
        {
            lock (_columnIndex)
            {
                int colPos = GetColumnPosition(column);
                colPos = SetValueColumn(row, column, value, colPos);
            }
        }

        private int SetValueColumn(int row, int column, T value, int colPos)
        {
            short page = (short)(row >> CellStoreSettings._pageBits);
            if (colPos >= 0)    //Column Found
            {
                ColumnIndex<T>? col = _columnIndex[colPos];
                int pagePos = col.GetPagePosition(row);
                if (pagePos < 0)
                {
                    pagePos = ~pagePos;
                    if (pagePos - 1 < 0 || col._pages[pagePos - 1].IndexOffset + CellStoreSettings._pageSize - 1 < row)
                    {
                        CellStore<T>.AddPage(col, pagePos, page);
                    }
                    else
                    {
                        pagePos--;
                    }
                }
                else if (pagePos >= col.PageCount)
                {
                    CellStore<T>.AddPage(col, pagePos, page);
                }
                PageIndex? pageItem = col._pages[pagePos];

                short ix = (short)(row - pageItem.IndexOffset);
                int cellPos = ArrayUtil.OptimizedBinarySearch(pageItem.Rows, ix, pageItem.RowCount);
                if (cellPos < 0)
                {
                    cellPos = ~cellPos;
                    AddCell(col, pagePos, cellPos, ix, value);
                }
                else
                {
                    col._values[pageItem.Rows[cellPos].IndexPointer] = value;
                }
            }
            else //Column does not exist
            {
                colPos = ~colPos;
                AddColumn(colPos, column);
                ColumnIndex<T>? col = _columnIndex[colPos];
                CellStore<T>.AddPage(col, 0, page);
                short ix = (short)(row - (page << CellStoreSettings._pageBits));
                AddCell(col, 0, 0, ix, value);
            }

            return colPos;
        }

        internal void Insert(int fromRow, int fromCol, int rows, int columns)
        {
            lock (_columnIndex)
            {
                if (ColumnCount == 0)
                {
                    return;
                }

                if (rows == 0)
                {
                    if (columns <= 0)
                    {
                        return;
                    }

                    //Entire column
                    int col = GetColumnPosition(fromCol);
                    if (col < 0)
                    {
                        col = ~col;
                    }
                    for (int c = col; c < ColumnCount; c++)
                    {
                        _columnIndex[c].Index += (short)columns;
                    }
                }
                else
                {
                    GetColumnPositionFromColumn(fromCol, columns, out int fromColPos, out int toColPos);

                    for (int c = fromColPos; c <= toColPos; c++)
                    {
                        ColumnIndex<T>? column = _columnIndex[c];
                        int pagePos = column.GetPagePosition(fromRow);
                        if (pagePos >= 0)
                        {
                            if (IsWithinPage(fromRow, column, pagePos)) //The row is inside the page
                            {
                                int rowPos = column._pages[pagePos].GetRowPosition(fromRow);
                                if (rowPos < 0)
                                {
                                    rowPos = ~rowPos;
                                }
                                InsertRowIntoPage(column, pagePos, rowPos, fromRow, rows);
                            }
                            else if (pagePos > 0 && IsWithinPage(fromRow, column, pagePos - 1)) //The row is inside the previous page
                            {
                                int rowPos = column._pages[pagePos - 1].GetRowPosition(fromRow);
                                if (rowPos > 0 && pagePos > 0)
                                {
                                    InsertRowIntoPage(column, pagePos - 1, rowPos, fromRow, rows);
                                }
                            }
                            else if (column.PageCount >= pagePos + 1)
                            {
                                int rowPos = column._pages[pagePos].GetRowPosition(fromRow);
                                if (rowPos < 0)
                                {
                                    rowPos = ~rowPos;
                                }
                                if (column._pages[pagePos].RowCount > rowPos)
                                {
                                    InsertRowIntoPage(column, pagePos, rowPos, fromRow, rows);
                                }
                                else
                                {
                                    InsertRowIntoPage(column, pagePos + 1, 0, fromRow, rows);
                                }
                            }
                        }
                        else
                        {
                            InsertRowIntoPage(column, ~pagePos, 0, fromRow, rows);
                        }
                    }
                }
            }
        }

        private void GetColumnPositionFromColumn(int fromCol, int columns, out int fromColPos, out int toColPos)
        {
            if (columns == 0)
            {
                fromColPos = 0;
                toColPos = ColumnCount - 1;
            }
            else
            {
                int endCol = fromCol + columns - 1;
                fromColPos = GetClosestColumnPosition(fromCol);
                toColPos = GetClosestColumnPosition(endCol);
                toColPos = Math.Min(toColPos, ColumnCount - 1);
                if (fromColPos < ColumnCount && _columnIndex[fromColPos].Index < fromCol)
                {
                    fromColPos++;
                }

                if (toColPos >= 0 && toColPos < ColumnCount && _columnIndex[toColPos].Index > endCol)
                {
                    toColPos--;
                }
            }
        }

        private static bool IsWithinPage(int row, ColumnIndex<T> column, int pagePos)
        {
            return (row >= column._pages[pagePos].MinIndex && row <= column._pages[pagePos].MaxIndex);
        }

        internal void Clear(int fromRow, int fromCol, int rows, int columns)
        {
            Delete(fromRow, fromCol, rows, columns, false);
        }
        internal void Delete(int fromRow, int fromCol, int rows, int columns)
        {
            Delete(fromRow, fromCol, rows, columns, true);
        }
        internal void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
        {
            lock (_columnIndex)
            {
                if (rows == 0)
                {
                    if (columns <= 0)
                    {
                        return;
                    }

                    DeleteColumns(fromCol, columns, shift);
                }
                else
                {
                    GetColumnPositionFromColumn(fromCol, columns, out int fromColPos, out int toColPos);

                    for (int c = fromColPos; c <= toColPos; c++)
                    {
                        ColumnIndex<T>? column = _columnIndex[c];
                        if (column.Index >= fromCol)
                        {
                            int toCol = fromCol + columns - 1;
                            if (column.Index > toCol)
                            {
                                break;
                            }

                            DeleteColumn(column, fromRow, rows, shift);
                        }
                    }
                }
            }
        }

        private static void DeleteColumn(ColumnIndex<T> column, int fromRow, int rows, bool shift)
        {
            int pagePos = column.GetPagePosition(fromRow);
            if (pagePos < 0)
            {
                pagePos = ~pagePos;
            }

            if (pagePos < column.PageCount)
            {
                int toRow = fromRow + rows - 1;
                PageIndex? page = column._pages[pagePos];
                if (page.StartsWithin(fromRow, toRow))
                {
                    //The deleted range starts within the page
                    pagePos = DeleteRows(column, pagePos, fromRow, rows, shift);
                }
                else if (column.PageCount > pagePos + 1)
                {
                    PageIndex? nextPage = column._pages[pagePos + 1];
                    if (nextPage.StartsWithin(fromRow, toRow))
                    {
                        pagePos = DeleteRows(column, pagePos + 1, fromRow, rows, shift);
                    }
                    else if (shift)
                    {
                        if (page.MaxIndex > toRow)
                        {
                            UpdatePageOffset(column, pagePos, -rows);
                        }
                        else
                        {
                            UpdatePageOffset(column, pagePos + 1, -rows);
                        }
                    }
                }
                else if (shift)
                {
                    if (pagePos < column.PageCount && column._pages[pagePos].MinIndex >= fromRow)
                    {
                        UpdatePageOffset(column, pagePos, -rows);
                    }
                }
            }
        }

        internal void DeleteShiftLeft(ExcelAddressBase fromAddress)
        {
            if (ColumnCount == 0)
            {
                return;
            }

            lock (_columnIndex)
            {
                short maxCol = _columnIndex[ColumnCount - 1].Index;
                int cols = fromAddress.Columns;
                for (int srcCol = fromAddress._toCol + 1; srcCol <= maxCol; srcCol++)
                {
                    int destCol = srcCol - cols;
                    MoveRangeColumnWise(srcCol, fromAddress._fromRow, fromAddress._toRow, destCol, fromAddress._fromRow);
                }
                Delete(fromAddress._fromRow, maxCol - cols + 1, fromAddress.Rows, cols, false);
            }
        }
        internal void InsertShiftRight(ExcelAddressBase fromAddress)
        {
            if (ColumnCount == 0)
            {
                return;
            }

            lock (_columnIndex)
            {
                short maxCol = _columnIndex[ColumnCount - 1].Index;
                for (int sourceCol = maxCol; sourceCol >= fromAddress._fromCol; sourceCol--)
                {
                    int destCol = fromAddress._toCol + 1 + (sourceCol - fromAddress._fromCol);
                    MoveRangeColumnWise(sourceCol, fromAddress._fromRow, fromAddress._toRow, destCol, fromAddress._fromRow);
                }
                Delete(fromAddress._fromRow, fromAddress._fromCol, fromAddress.Rows, fromAddress.Columns, false);
            }
        }

        private void MoveRangeColumnWise(int sourceCol, int sourceStartRow, int sourceEndRow, int destCol, int destStartRow)
        {
            int sourceColPos = GetColumnPosition(sourceCol);
            int destColPos = GetColumnPosition(destCol);

            int rows = sourceEndRow - sourceStartRow + 1;
            if (sourceColPos < 0 && destColPos < 0)             //Neither source nor destiontion exists, so we're done
            {
                return;
            }
            else if (sourceColPos < 0 && destCol >= 0)          //Source column does not exist, delete range in destintation column
            {
                Delete(destStartRow, destCol, rows, 1, false);
                return;
            }

            ColumnIndex<T>? sourceColIx = _columnIndex[sourceColPos];
            int sourcePagePos = sourceColIx.GetPagePosition(sourceStartRow);
            if (sourcePagePos < 0)
            {
                sourcePagePos = ~sourcePagePos;
            }
            if (sourcePagePos > sourceColIx._pages.Length - 1 || sourceColIx.PageCount == 0)
            {
                Delete(destStartRow, destCol, rows, 1, false);
                return;
            }

            PageIndex? sourcePage = sourceColIx._pages[sourcePagePos];

            int sourceRowIx = sourcePage.GetRowPosition(sourceStartRow);
            if (sourceRowIx < 0)
            {
                sourceRowIx = ~sourceRowIx;
                if (sourceRowIx >= sourcePage.RowCount || sourcePage.GetRow(sourceRowIx) < sourceStartRow)
                {
                    Delete(destStartRow, destCol, rows, 1, false);
                    return;
                }
            }

            //Get and create the destination column
            ColumnIndex<T> destColIx;
            if (destColPos < 0 && sourceRowIx >= 0 && sourcePage.GetRow(sourceRowIx) <= sourceEndRow)
            {
                destColPos = ~destColPos;
                AddColumn(destColPos, destCol);
                destColIx = _columnIndex[destColPos];
            }
            else if (destColPos >= 0)
            {
                destColIx = _columnIndex[destColPos];
                //No rows to move, just clear the destination
                DeleteColumn(destColIx, destStartRow, rows, false);
            }
            else
            {
                return;
            }

            if (sourcePage.GetRow(sourceRowIx) > sourceEndRow)
            {
                return;
            }

            //Start copy
            int sourceRow = sourcePage.GetRow(sourceRowIx);
            int prevDestPagePos = -1;
            int destRowIx = -1;
            do
            {
                int destRow = destStartRow + (sourceRow - sourceStartRow);
                int destPagePos = destColIx.GetPagePosition(destRow);
                if (destPagePos < 0)
                {
                    destPagePos = ~destPagePos;
                    short page = (short)(destRow >> CellStoreSettings._pageBits);
                    CellStore<T>.AddPage(destColIx, destPagePos, page);
                }
                PageIndex? destPage = destColIx._pages[destPagePos];
                destColIx._values.Add(sourceColIx._values[sourcePage.Rows[sourceRowIx].IndexPointer]);
                if (prevDestPagePos == destPagePos)
                {
                    AddCellPointer(destColIx, ref destPagePos, ref destRowIx, (short)(destRow - destPage.IndexOffset), destColIx._values.Count - 1);
                }
                else
                {
                    if (destRowIx == -1)
                    {
                        destRowIx = ~destColIx._pages[destPagePos].GetRowPosition(destRow);
                    }
                    else
                    {
                        destRowIx = 0;
                    }

                    AddCellPointer(destColIx, ref destPagePos, ref destRowIx, (short)(destRow - destPage.IndexOffset), destColIx._values.Count - 1);
                }
                sourceRowIx++;
                destRowIx++;
                if (sourceRowIx == sourcePage.RowCount)
                {
                    sourcePagePos++;
                    if (sourcePagePos >= sourceColIx.PageCount)
                    {
                        break;
                    }

                    sourcePage = sourceColIx._pages[sourcePagePos];
                    sourceRowIx = 0;
                }
                sourceRow = sourcePage.GetRow(sourceRowIx);
                prevDestPagePos = destPagePos;
            }
            while (sourceRow <= sourceEndRow);
        }

        /// <summary>
        /// Delete a number of rows from a specific row
        /// </summary>
        /// <param name="fromRow">The first row to delete</param>
        /// <param name="rows">Number of rows</param>
        /// <param name="shift">If rows are shifted upwards</param>
        /// <param name="column">The column index</param>
        /// <param name="pagePos">The page position</param>
        /// <returns></returns>
        private static int DeleteRows(ColumnIndex<T> column, int pagePos, int fromRow, int rows, bool shift)
        {
            int toRow = fromRow + rows - 1;
            PageIndex? page = column._pages[pagePos];
            int rowsLeft = rows;

            if (!page.IsWithin(fromRow, toRow))
            {
                //DeleteCells
                rowsLeft = DeleteRowsInsidePage(column, pagePos, fromRow, toRow, shift);
                if (rowsLeft > 0)
                {
                    pagePos++;
                }
            }

            if (rowsLeft > 0 && pagePos < column.PageCount && column._pages[pagePos].MinIndex <= toRow)
            {
                int delFromRow = shift ? fromRow : toRow - rowsLeft + 1;
                rowsLeft = DeletePages(delFromRow, rowsLeft, column, pagePos, shift);
                if (rowsLeft > 0)
                {
                    delFromRow = shift ? fromRow : toRow - rowsLeft + 1;
                    pagePos = column.GetPagePosition(delFromRow);
                    if (pagePos < 0)
                    {
                        pagePos = ~pagePos;
                    }

                    DeleteRowsInsidePage(column, pagePos, delFromRow, shift ? delFromRow + rowsLeft - 1 : toRow, shift);
                }
            }

            return pagePos;
        }
        private static void UpdatePageOffset(ColumnIndex<T> column, int pagePos, int rows)
        {
            //Update Pageoffsets

            if (pagePos < column.PageCount)
            {
                for (int p = pagePos; p < column.PageCount; p++)
                {
                    p = CellStore<T>.UpdatePageOffsetSinglePage(column, p, rows);
                }
            }
        }

        private static int UpdatePageOffsetSinglePage(ColumnIndex<T> column, int pagePos, int rows)
        {
            PageIndex? page = column._pages[pagePos];
            if (Math.Abs(page.Offset + rows) < CellStoreSettings._pageSize)
            {
                page.Offset += rows;
            }
            else
            {
                short indexAdd = (short)((page.Offset + rows) / CellStoreSettings._pageSize);
                page.Index += indexAdd;
                page.Offset = (page.Offset + rows) % CellStoreSettings._pageSize;

                //Verify if merge should be done.
                if (pagePos > 0 && page.Index == column._pages[pagePos - 1].Index)
                {
                    CellStore<T>.MergePage(column, pagePos - 1);
                    return pagePos - 1;
                }
            }

            return pagePos;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fromRow">From row</param>
        /// <param name="rows">Number of rows</param>
        /// <param name="column">The column index</param>
        /// <param name="pagePos">The page position</param>
        /// <param name="shift">Shift cells or not</param>
        /// <returns>Return rows left to delete, for DeleteCells</returns>
        private static int DeletePages(int fromRow, int rows, ColumnIndex<T> column, int pagePos, bool shift)
        {
            PageIndex page = column._pages[pagePos];
            int pageStartRow = fromRow;
            int startRows = rows;
            while (page != null && page.MinIndex >= fromRow &&
                    ((shift && page.MaxIndex < fromRow + rows) ||
                    (!shift && page.MaxIndex < fromRow + startRows)))
            {
                //Delete entire page.
                int delSize = page.MaxIndex - pageStartRow + 1;
                int prevMaxIndex = page.MaxIndex;
                rows -= delSize;
                int prevOffset = page.Offset;
                Array.Copy(column._pages, pagePos + 1, column._pages, pagePos, column.PageCount - pagePos - 1);
                column.PageCount--;

                if (column.PageCount == 0)
                {
                    return 0;
                }
                if (shift)
                {
                    UpdatePageOffset(column, pagePos, -delSize);
                }
                if (column.PageCount > pagePos)
                {
                    if (shift)
                    {
                        if (pagePos == 0)
                        {
                            pageStartRow = 1;
                        }
                        else
                        {
                            pageStartRow = column._pages[pagePos - 1].MaxIndex + 1;
                        }
                    }
                    else
                    {
                        pageStartRow = prevMaxIndex + 1;
                    }
                    page = column._pages[pagePos];
                }
                else
                {
                    //No more pages, return 0
                    return 0;
                }
            }
            return rows;
        }
        ///
        private static int DeleteRowsInsidePage(ColumnIndex<T> column, int pagePos, int fromRow, int toRow, bool shift)
        {
            PageIndex? page = column._pages[pagePos];
            int deletedRows = 0;

            int fromPos = page.GetRowPosition(fromRow);
            if (fromPos < 0)
            {
                fromPos = ~fromPos;
            }
            int toPos = page.GetRowPosition(toRow);
            if (toPos < 0)
            {
                toPos = ~toPos - 1;
            }

            if (fromPos < page.RowCount)
            {
                int maxRow = page.MaxIndex;
                if (toRow >= maxRow)
                {
                    if (fromRow == page.MinIndex) //Delete entire page, TODO: Remove when tests a good,
                    {
                        throw (new Exception("Invalid cell delete: DeleteCells"));
                    }
                    page.RowCount -= page.RowCount - fromPos;
                    deletedRows = maxRow - fromRow + 1;
                }
                else
                {
                    deletedRows = toRow - fromRow + 1;
                    if (fromPos <= toPos)
                    {
                        Array.Copy(page.Rows, toPos + 1, page.Rows, fromPos, page.RowCount - toPos - 1);
                        page.RowCount -= toPos - fromPos + 1;
                    }
                    if (shift && fromPos > 0)
                    {
                        //If the page is not updated from start, we must update the row indexs. Otherwise we will update the whole page in the UpdatePageOffset futher down.
                        UpdateRowIndex(page, fromPos, deletedRows);
                    }
                }
            }
            else if (shift)
            {
                UpdateRowIndex(page, toPos, toRow - fromRow);
            }

            if (shift && deletedRows > 0)
            {
                if (fromPos > 0)
                {
                    pagePos++;
                }
                UpdatePageOffset(column, pagePos, -deletedRows);
            }
            return (toRow - fromRow + 1) - deletedRows;
        }

        private static void UpdateRowIndex(PageIndex page, int toPos, int rows)
        {
            for (int r = toPos; r < page.RowCount; r++)
            {
                page.Rows[r].Index -= (short)rows;
            }
        }

        private void DeleteColumns(int fromCol, int columns, bool shift)
        {
            int fPos = GetColumnPosition(fromCol);
            if (fPos < 0)
            {
                fPos = ~fPos;
            }
            int tPos = fPos;
            for (int c = fPos; c <= ColumnCount; c++)
            {
                tPos = c;
                if (tPos == ColumnCount || _columnIndex[c].Index >= fromCol + columns)
                {
                    break;
                }
            }

            if (ColumnCount <= fPos)
            {
                return;
            }

            if (_columnIndex[fPos].Index >= fromCol && _columnIndex[fPos].Index <= fromCol + columns)
            {
                if (tPos < ColumnCount)
                {
                    Array.Copy(_columnIndex, tPos, _columnIndex, fPos, ColumnCount - tPos);
                }
                ColumnCount -= (tPos - fPos);
            }
            if (shift)
            {
                for (int c = fPos; c < ColumnCount; c++)
                {
                    _columnIndex[c].Index -= (short)columns;
                }
            }
        }

        private static void InsertRowIntoPage(ColumnIndex<T> column, int pagePos, int rowPos, int row, int rows)
        {
            if (pagePos >= column.PageCount)
            {
                return;    //A page after last cell.
            }

            PageIndex? page = column._pages[pagePos];
            if (rowPos > 0) //RowPos is 0 then we can update the page index instead
            {
                if (rows >= CellStoreSettings._pageSize)
                {
                    SplitPageAtPosition(column, pagePos, page, rowPos);
                    CellStore<T>.UpdatePageOffsetSinglePage(column, ++pagePos, rows);
                }
                else
                {
                    CellStore<T>.AddRowIndex(rowPos, (short)rows, page);
                    pagePos = ValidateAndSplitPageIfNeeded(column, page, pagePos);
                }
            }
            else
            {
                pagePos--; // We want to adjust the offset of the current page as well, as rowPos == 0
            }

            UpdatePageOffset(column, pagePos + 1, rows);
        }

        private static int ValidateAndSplitPageIfNeeded(ColumnIndex<T> column, PageIndex page, int pagePos)
        {
            if (page.RowSpan >= CellStoreSettings._pageSizeMax)   //Cannot be larger than the max size of the page.
            {
                pagePos = SplitPage(column, pagePos);
            }

            return pagePos;
        }

        private static void AddRowIndex(int rowPos, short rows, PageIndex page)
        {
            //Add to Pages.
            for (int r = rowPos; r < page.RowCount; r++)
            {
                page.Rows[r].Index += rows;
            }
        }
        private static void MergePage(ColumnIndex<T> column, int pagePos)
        {
            PageIndex? Page1 = column._pages[pagePos];
            PageIndex? Page2 = column._pages[pagePos + 1];

            PageIndex? newPage = new PageIndex(Page1, 0, Page1.RowCount + Page2.RowCount);
            newPage.RowCount = Page1.RowCount + Page2.RowCount;
            Array.Copy(Page1.Rows, 0, newPage.Rows, 0, Page1.RowCount);
            Array.Copy(Page2.Rows, 0, newPage.Rows, Page1.RowCount, Page2.RowCount);
            for (int r = Page1.RowCount; r < newPage.RowCount; r++)
            {
                newPage.Rows[r].Index += (short)(Page2.IndexOffset - Page1.IndexOffset);
            }

            column._pages[pagePos] = newPage;
            column.PageCount--;

            if (column.PageCount > (pagePos + 1))
            {
                Array.Copy(column._pages, pagePos + 2, column._pages, pagePos + 1, column.PageCount - (pagePos + 1));
            }
        }

        internal static int GetSize(int size)
        {
            int newSize = 16;
            while (newSize < size)
            {
                newSize <<= 1;
            }
            return newSize;
        }
        private static void AddCell(ColumnIndex<T> columnIndex, int pagePos, int pos, short ix, T value)
        {
            PageIndex pageItem = MakeRoomInPage(columnIndex, ref pagePos, ref pos);
            pageItem.Rows[pos] = new IndexItem() { Index = ix, IndexPointer = columnIndex._values.Count };
            columnIndex._values.Add(value);
            pageItem.RowCount++;
        }
        private static void AddCellPointer(ColumnIndex<T> columnIndex, ref int pagePos, ref int pos, short ix, int pointer)
        {
            PageIndex pageItem = MakeRoomInPage(columnIndex, ref pagePos, ref pos);
            pageItem.Rows[pos] = new IndexItem() { Index = ix, IndexPointer = pointer };
            pageItem.RowCount++;
        }
        private static PageIndex MakeRoomInPage(ColumnIndex<T> columnIndex, ref int pagePos, ref int pos)
        {
            PageIndex? pageItem = columnIndex._pages[pagePos];
            if (pageItem.RowCount == pageItem.Rows.Length)
            {
                if (pageItem.RowCount == CellStoreSettings._pageSizeMax) //Max size-->Split
                {
                    pagePos = SplitPage(columnIndex, pagePos);
                    if (columnIndex._pages[pagePos - 1].RowCount > pos)
                    {
                        pagePos--;
                    }
                    else
                    {
                        pos -= columnIndex._pages[pagePos - 1].RowCount;
                    }
                    pageItem = columnIndex._pages[pagePos];
                }
                else //Expand to double size.
                {
                    IndexItem[]? rowsTmp = new IndexItem[pageItem.Rows.Length << 1];
                    Array.Copy(pageItem.Rows, 0, rowsTmp, 0, pageItem.RowCount);
                    pageItem.Rows = rowsTmp;
                }
            }
            if (pos < pageItem.RowCount)
            {
                Array.Copy(pageItem.Rows, pos, pageItem.Rows, pos + 1, pageItem.RowCount - pos);
            }
            return pageItem;
        }

        private static int SplitPage(ColumnIndex<T> columnIndex, int pagePos)
        {
            PageIndex? page = columnIndex._pages[pagePos];
            ResetPageOffset(page);

            //Find split position
            int splitPos = ArrayUtil.OptimizedBinarySearch(page.Rows, CellStoreSettings._pageSize, page.RowCount);
            if (splitPos < 0)
            {
                splitPos = ~splitPos;
            }

            SplitPageAtPosition(columnIndex, pagePos, page, splitPos);
            return pagePos + 1;
        }

        private static void SplitPageAtPosition(ColumnIndex<T> columnIndex, int pagePos, PageIndex page, int splitPos)
        {
            PageIndex? nextPage = new PageIndex(page, splitPos, page.RowCount - splitPos, (short)(page.Index + 1), page.Offset, CellStoreSettings._pageSizeMax);
            page.RowCount = splitPos;

            for (int r = 0; r < nextPage.RowCount; r++)
            {
                nextPage.Rows[r].Index = (short)(nextPage.Rows[r].Index - CellStoreSettings._pageSize);
            }

            CellStore<T>.AddPage(columnIndex, nextPage, pagePos + 1);
        }

        private static void ResizePageCollectionIfNecessery(ColumnIndex<T> columnIndex)
        {
            if (columnIndex.PageCount >= columnIndex._pages.Length)
            {
                PageIndex[]? pageTmp = new PageIndex[columnIndex._pages.Length << 1];    //Double size
                Array.Copy(columnIndex._pages, 0, pageTmp, 0, columnIndex.PageCount);
                columnIndex._pages = pageTmp;
            }
        }

        private static void ResetPageOffset(PageIndex page)
        {
            if (page.Offset != 0)
            {
                int offset = page.Offset;
                page.Offset = 0;
                for (int r = 0; r < page.RowCount; r++)
                {
                    page.Rows[r].Index += (short)offset;
                }
            }
        }

        private static void AddPage(ColumnIndex<T> column, int pos, short index)
        {
            CellStore<T>.AddPage(column, pos);
            column._pages[pos] = new PageIndex(CellStoreSettings._pageSizeMin) { Index = index };
            if (pos > 0)
            {
                PageIndex? pp = column._pages[pos - 1];
                if (pp.RowCount > 0 && pp.Rows[pp.RowCount - 1].Index > CellStoreSettings._pageSize)
                {
                    column._pages[pos].Offset = pp.Rows[pp.RowCount - 1].Index - CellStoreSettings._pageSize;
                }
            }
        }
        /// <summary>
        /// Add a new page to the collection
        /// </summary>
        /// <param name="column">The column</param>
        /// <param name="pos">Position</param>
        /// <param name="page">The new page object to add</param>
        private static void AddPage(ColumnIndex<T> column, PageIndex page, int pos)
        {
            CellStore<T>.AddPage(column, pos);
            column._pages[pos] = page;
        }
        /// <summary>
        /// Add a new page to the collection
        /// </summary>
        /// <param name="column">The column</param>
        /// <param name="pos">Position</param>
        private static void AddPage(ColumnIndex<T> column, int pos)
        {
            ResizePageCollectionIfNecessery(column);

            if (pos < column.PageCount)
            {
                Array.Copy(column._pages, pos, column._pages, pos + 1, column.PageCount - pos);
            }
            column.PageCount++;
        }
        private void AddColumn(int pos, int Column)
        {
            if (ColumnCount == _columnIndex.Length)
            {
                ColumnIndex<T>[]? colTmp = new ColumnIndex<T>[_columnIndex.Length * 2];
                Array.Copy(_columnIndex, 0, colTmp, 0, ColumnCount);
                _columnIndex = colTmp;
            }
            if (pos < ColumnCount)
            {
                Array.Copy(_columnIndex, pos, _columnIndex, pos + 1, ColumnCount - pos);
            }
            _columnIndex[pos] = new ColumnIndex<T>() { Index = (short)(Column) };
            ColumnCount++;
        }

        public void Dispose()
        {
            if (_columnIndex == null)
            {
                return;
            }

            for (int c = 0; c < ColumnCount; c++)
            {
                if (_columnIndex[c] != null)
                {
                    ((IDisposable)_columnIndex[c]).Dispose();
                }
            }
            _columnIndex = null;
        }

        internal bool NextCell(ref int row, ref int col)
        {

            return NextCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        internal bool NextCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
        {
            if (minColPos >= ColumnCount)
            {
                return false;
            }
            if (maxColPos >= ColumnCount)
            {
                maxColPos = ColumnCount - 1;
            }
            int c = GetColumnPosition(col);
            if (c >= 0)
            {
                if (c > maxColPos)
                {
                    if (col <= minColPos)
                    {
                        return false;
                    }
                    col = minColPos;
                    return NextCell(ref row, ref col);
                }
                else
                {
                    bool r = GetNextCell(ref row, ref c, minColPos, maxRow, maxColPos);
                    col = _columnIndex[c].Index;
                    return r;
                }
            }
            else
            {
                c = ~c;
                if (c >= ColumnCount)
                {
                    c = this.ColumnCount - 1;
                }

                if (col > _columnIndex[c].Index)
                {
                    if (col <= minColPos)
                    {
                        return false;
                    }
                    col = minColPos;
                    return NextCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    bool r = GetNextCell(ref row, ref c, minColPos, maxRow, maxColPos);
                    if (r)
                    {
                        col = _columnIndex[c].Index;
                    }
                    return r;
                }
            }
        }
        internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos)
        {
            if (ColumnCount == 0)
            {
                return false;
            }
            else
            {
                if (++colPos < ColumnCount && colPos <= endColPos)
                {
                    int r = _columnIndex[colPos].GetNextRow(row);
                    if (r == row) //Exists next Row
                    {
                        return true;
                    }
                    else
                    {
                        int minRow, minCol;
                        if (r > row)
                        {
                            minRow = r;
                            minCol = colPos;
                        }
                        else
                        {
                            minRow = int.MaxValue;
                            minCol = 0;
                        }

                        int c = colPos + 1;
                        while (c < ColumnCount && c <= endColPos)
                        {
                            r = _columnIndex[c].GetNextRow(row);
                            if (r == row) //Exists next Row
                            {
                                colPos = c;
                                return true;
                            }
                            if (r > row && r < minRow)
                            {
                                minRow = r;
                                minCol = c;
                            }
                            c++;
                        }
                        c = startColPos;
                        if (row < endRow)
                        {
                            row++;
                            while (c < colPos)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && (r < minRow || (r == minRow && c < minCol)) && r <= endRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c++;
                            }
                        }

                        if (minRow == int.MaxValue || minRow > endRow)
                        {
                            return false;
                        }
                        else
                        {
                            row = minRow;
                            colPos = minCol;
                            return true;
                        }
                    }
                }
                else
                {
                    if (colPos <= startColPos || row >= endRow)
                    {
                        return false;
                    }
                    colPos = startColPos - 1;
                    row++;
                    return GetNextCell(ref row, ref colPos, startColPos, endRow, endColPos);
                }
            }
        }
        internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos, ref int[] pagePos, ref int[] cellPos)
        {
            if (colPos == endColPos)
            {
                colPos = startColPos;
                row++;
            }
            else
            {
                colPos++;
            }

            if (pagePos[colPos] < 0)
            {
                if (pagePos[colPos] == -1)
                {
                    pagePos[colPos] = _columnIndex[colPos].GetPagePosition(row);
                }
            }
            else if (_columnIndex[colPos]._pages[pagePos[colPos]].RowCount <= row)
            {
                if (_columnIndex[colPos].PageCount > pagePos[colPos])
                {
                    pagePos[colPos]++;
                }
                else
                {
                    pagePos[colPos] = -2;
                }
            }

            int r = _columnIndex[colPos]._pages[pagePos[colPos]].IndexOffset + _columnIndex[colPos]._pages[pagePos[colPos]].Rows[cellPos[colPos]].Index;
            if (r == row)
            {
                row = r;
            }
            else
            {
            }
            return true;
        }
        internal bool PrevCell(ref int row, ref int col)
        {
            return PrevCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        internal bool PrevCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
        {
            if (minColPos >= ColumnCount)
            {
                return false;
            }
            if (maxColPos >= ColumnCount)
            {
                maxColPos = ColumnCount - 1;
            }
            int c = GetColumnPosition(col);
            if (c >= 0)
            {
                if (c == 0)
                {
                    if (col >= maxColPos)
                    {
                        return false;
                    }
                    if (row == minRow)
                    {
                        return false;
                    }
                    row--;
                    col = maxColPos;
                    return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    bool ret = GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
                    if (ret)
                    {
                        col = _columnIndex[c].Index;
                    }
                    return ret;
                }
            }
            else
            {
                c = ~c;
                if (c == 0)
                {
                    if (col >= maxColPos || row <= 0)
                    {
                        return false;
                    }
                    col = maxColPos;
                    row--;
                    return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    bool ret = GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
                    if (ret)
                    {
                        col = _columnIndex[c].Index;
                    }
                    return ret;
                }
            }
        }
        internal bool GetPrevCell(ref int row, ref int colPos, int startRow, int startColPos, int endColPos)
        {
            if (ColumnCount == 0)
            {
                return false;
            }
            else
            {
                if (--colPos >= startColPos)
                {
                    int r = _columnIndex[colPos].GetNextRow(row);
                    if (r == row) //Exists next Row
                    {
                        return true;
                    }
                    else
                    {
                        int minRow, minCol;
                        if (r > row && r >= startRow)
                        {
                            minRow = r;
                            minCol = colPos;
                        }
                        else
                        {
                            minRow = int.MaxValue;
                            minCol = 0;
                        }

                        int c = colPos - 1;
                        if (c >= startColPos)
                        {
                            while (c >= startColPos)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && r < minRow && r >= startRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c--;
                            }
                        }
                        if (row > startRow)
                        {
                            c = endColPos;
                            row--;
                            while (c > colPos)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && r < minRow && r >= startRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c--;
                            }
                        }
                        if (minRow == int.MaxValue || startRow < minRow)
                        {
                            return false;
                        }
                        else
                        {
                            row = minRow;
                            colPos = minCol;
                            return true;
                        }
                    }
                }
                else
                {
                    colPos = ColumnCount;
                    row--;
                    if (row < startRow)
                    {
                        return false;
                    }
                    else
                    {
                        return GetPrevCell(ref colPos, ref row, startRow, startColPos, endColPos);
                    }
                }
            }
        }
        /// <summary>
        /// Before enumerating columns where values are set to the cells store, 
        /// this method makes sure the columns are created before the enumerator is created, so the positions will not get out of sync when a new column is added.
        /// </summary>
        /// <param name="fromCol">From column</param>
        /// <param name="toCol">To Column</param>
        internal void EnsureColumnsExists(int fromCol, int toCol)
        {
            for (int col = fromCol; col <= toCol; col++)
            {
                int colPos = GetColumnPosition(col);
                if (colPos < 0)
                {
                    colPos = ~colPos;
                    AddColumn(colPos, col);
                }
            }
        }
    }
}