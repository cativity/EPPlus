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
using System.Collections.Generic;
using System.Collections;

namespace OfficeOpenXml.Core.CellStore
{
    internal class CellStoreEnumerator<T> : IEnumerable<T>, IEnumerator<T>
    {
        CellStore<T> _cellStore;
        int row, colPos;
        int[] pagePos, cellPos;
        int _startRow, _startCol, _endRow, _endCol;
        int minRow, minColPos, maxRow, maxColPos;
        public CellStoreEnumerator(CellStore<T> cellStore) :
            this(cellStore, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns)
        {
        }
        public CellStoreEnumerator(CellStore<T> cellStore, int StartRow, int StartCol, int EndRow, int EndCol)
        {
            this._cellStore = cellStore;

            this._startRow = StartRow;
            this._startCol = StartCol;
            this._endRow = EndRow;
            this._endCol = EndCol;

            this.Init();

        }

        internal void Init()
        {
            this.minRow = this._startRow;
            this.maxRow = this._endRow;

            this.minColPos = this._cellStore.GetColumnPosition(this._startCol);
            if (this.minColPos < 0)
            {
                this.minColPos = ~this.minColPos;
            }

            this.maxColPos = this._cellStore.GetColumnPosition(this._endCol);
            if (this.maxColPos < 0)
            {
                this.maxColPos = ~this.maxColPos - 1;
            }

            this.row = this.minRow;
            this.colPos = this.minColPos - 1;

            int cols = this.maxColPos - this.minColPos + 1;
            this.pagePos = new int[cols];
            this.cellPos = new int[cols];
            for (int i = 0; i < cols; i++)
            {
                this.pagePos[i] = -1;
                this.cellPos[i] = -1;
            }
        }
        internal int Row
        {
            get
            {
                return this.row;
            }
        }
        internal int Column
        {
            get
            {
                if (this.colPos<0 || this.colPos>= this._cellStore.ColumnCount)
                {
                    return -1;
                }
                return this._cellStore._columnIndex[this.colPos].Index;
            }
        }
        internal T Value
        {
            get
            {
                lock (this._cellStore)
                {
                    return this._cellStore.GetValue(this.row, this.Column);
                }
            }
            set
            {
                lock (this._cellStore)
                {
                    this._cellStore.SetValue(this.row, this.Column, value);
                }
            }
        }
        internal bool Next()
        {
            return this._cellStore.GetNextCell(ref this.row, ref this.colPos, this.minColPos, this.maxRow, this.maxColPos);
        }
        internal bool Previous()
        {
            lock (this._cellStore)
            {
                return this._cellStore.GetPrevCell(ref this.row, ref this.colPos, this.minRow, this.minColPos, this.maxColPos);
            }
        }

        public string CellAddress
        {
            get
            {
                return ExcelCellBase.GetAddress(this.Row, this.Column);
            }
        }

        public IEnumerator<T> GetEnumerator()
        {
            this.Reset();
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            this.Reset();
            return this;
        }

        public T Current
        {
            get
            {
                return this.Value;
            }
        }

        public void Dispose()
        {

        }

        object IEnumerator.Current
        {
            get
            {
                this.Reset();
                return this;
            }
        }

        public bool MoveNext()
        {
            return this.Next();
        }

        public void Reset()
        {
            this.Init();
        }
    }
}