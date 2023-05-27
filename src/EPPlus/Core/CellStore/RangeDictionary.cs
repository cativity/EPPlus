using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System.Runtime.CompilerServices;
using System.Net;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Security.Cryptography;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.Core.CellStore;

/// <summary>
/// This class stores ranges to keep track if they have been accessed before and adds a reference to <see cref="T"/>.
/// </summary>
internal class RangeDictionary<T>
{
    internal struct RangeItem : IComparable<RangeItem>
    {
        public RangeItem(long rowSpan)
        {
            this.RowSpan = rowSpan;
            this.Value = default;
        }

        public RangeItem(long rowSpan, T value)
        {
            this.RowSpan = rowSpan;
            this.Value = value;
        }

        internal long RowSpan;
        internal T Value;

        public int CompareTo(RangeItem other)
        {
            return this.RowSpan.CompareTo(other.RowSpan);
        }

        public override string ToString()
        {
            int fr = (int)(this.RowSpan >> 20) + 1;
            int tr = (int)(this.RowSpan & 0xFFFFF) + 1;

            return $"{fr} - {tr}";
        }
    }

    internal Dictionary<int, List<RangeItem>> _addresses = new Dictionary<int, List<RangeItem>>();
    private bool _extendValuesToInsertedColumn = true;

    internal bool Exists(int fromRow, int fromCol, int toRow, int toCol)
    {
        for (int c = fromCol; c <= toCol; c++)
        {
            long rowSpan = (((long)fromRow - 1) << 20) | ((long)toRow - 1);
            RangeItem ri = new RangeItem(rowSpan, default);

            if (this._addresses.TryGetValue(c, out List<RangeItem> rows))
            {
                int ix = rows.BinarySearch(ri);

                if (ix >= 0)
                {
                    return true;
                }
                else if (rows.Count > 0)
                {
                    ix = ~ix;

                    if (ix < rows.Count && ExistsInSpan(fromRow, toRow, rows[ix].RowSpan))
                    {
                        return true;
                    }
                    else if (--ix < rows.Count && ix >= 0)
                    {
                        return ExistsInSpan(fromRow, toRow, rows[ix].RowSpan);
                    }
                }
            }
        }

        return false;
    }

    internal bool Exists(int row, int col)
    {
        if (this._addresses.TryGetValue(col, out List<RangeItem> rows))
        {
            long rowSpan = ((row - 1) << 20) | (row - 1);
            RangeItem ri = new RangeItem(rowSpan, default);
            int ix = rows.BinarySearch(ri);

            if (ix < 0)
            {
                ix = ~ix;

                if (ix < rows.Count)
                {
                    if (ExistsInSpan(row, row, rows[ix].RowSpan))
                    {
                        return true;
                    }
                }

                if (ix > 0 && --ix < rows.Count)
                {
                    return ExistsInSpan(row, row, rows[ix].RowSpan);
                }
            }
            else
            {
                return true;
            }
        }

        return false;
    }

    internal T this[int row, int column]
    {
        get
        {
            if (this._addresses.TryGetValue(column, out List<RangeItem> rows))
            {
                long rowSpan = ((row - 1) << 20) | (row - 1);
                RangeItem ri = new RangeItem(rowSpan, default);
                int ix = rows.BinarySearch(ri);

                if (ix < 0)
                {
                    ix = ~ix;

                    if (ix < rows.Count)
                    {
                        if (ExistsInSpan(row, row, rows[ix].RowSpan))
                        {
                            return rows[ix].Value;
                        }
                    }

                    if (--ix < rows.Count && ix >= 0)
                    {
                        if (ExistsInSpan(row, row, rows[ix].RowSpan))
                        {
                            return rows[ix].Value;
                        }
                    }
                }
                else
                {
                    return rows[ix].Value;
                }
            }

            return default;
        }
    }

    internal List<T> GetValuesFromRange(int fromRow, int fromCol, int toRow, int toCol)
    {
        HashSet<T>? hs = new HashSet<T>();
        long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);
        RangeItem searchItem = new RangeItem(rowSpan, default);
        int minCol = this._addresses.Keys.Min();
        int maxCol = this._addresses.Keys.Max();
        fromCol = fromCol < minCol ? minCol : fromCol;

        for (int col = fromCol; col <= toCol; col++)
        {
            if (col > maxCol)
            {
                break;
            }

            if (this._addresses.TryGetValue(col, out List<RangeItem> rows))
            {
                int ix = rows.BinarySearch(searchItem);

                if (ix < 0)
                {
                    ix = ~ix;

                    if (ix > 0)
                    {
                        ix--;
                    }
                }

                while (ix < rows.Count)
                {
                    RangeItem ri = rows[ix];
                    int fr = (int)(ri.RowSpan >> 20) + 1;
                    int tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                    if (tr < fromRow)
                    {
                        ix++;

                        continue;
                    }

                    if (fromRow <= tr && toRow >= fr)
                    {
                        if (!hs.Contains(ri.Value))
                        {
                            hs.Add(ri.Value);
                        }

                        ix++;
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }

        return hs.ToList();
    }

    internal void Merge(int fromRow, int fromCol, int toRow, int toCol, T value)
    {
        for (int c = fromCol; c <= toCol; c++)
        {
            this.MergeRowSpan(c, fromRow, toRow, value);
        }
    }

    internal void Add(int fromRow, int fromCol, int toRow, int toCol, T value)
    {
        if (this.Exists(fromRow, fromCol, toRow, toCol))
        {
            throw new InvalidOperationException($"Range already starting from range {ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol)}");
        }

        for (int c = fromCol; c <= toCol; c++)
        {
            this.AddRowSpan(c, fromRow, toRow, value);
        }
    }

    internal void Add(int row, int col, T value)
    {
        if (this.Exists(row, col))
        {
            throw new InvalidOperationException($"Range already starting from cell {ExcelCellBase.GetAddress(row, col)}");
        }

        this.AddRowSpan(col, row, row, value);
    }

    internal void InsertRow(int fromRow, int noRows, int fromCol = 1, int toCol = ExcelPackage.MaxColumns)
    {
        long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);

        foreach (int c in this._addresses.Keys)
        {
            if (c >= fromCol && c <= toCol)
            {
                List<RangeItem>? rows = this._addresses[c];
                RangeItem ri = new RangeItem(rowSpan, default);
                int ix = rows.BinarySearch(ri);

                if (ix < 0)
                {
                    ix = ~ix;

                    if (ix > 0)
                    {
                        ix--;
                    }
                }

                if (ix < rows.Count)
                {
                    ri = rows[ix];
                    int fr = (int)(ri.RowSpan >> 20) + 1;
                    int tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                    if (fr >= fromRow)
                    {
                        ri.RowSpan = ((fr + noRows - 1) << 20) | (tr + noRows - 1);
                    }
                    else
                    {
                        ri.RowSpan = ((fr - 1) << 20) | (tr + noRows - 1);
                    }

                    rows[ix] = ri;
                }

                int add = (noRows << 20) | noRows;

                for (int i = ix + 1; i < rows.Count; i++)
                {
                    rows[i] = new RangeItem(rows[i].RowSpan + add, rows[i].Value);
                }
            }
        }
    }

    internal void DeleteRow(int fromRow, int noRows, int fromCol = 1, int toCol = ExcelPackage.MaxColumns, bool shiftRow = true)
    {
        long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);

        foreach (int c in this._addresses.Keys)
        {
            if (c >= fromCol && c <= toCol)
            {
                List<RangeItem>? rows = this._addresses[c];
                RangeItem ri = new RangeItem(rowSpan);
                int rowStartIndex = rows.BinarySearch(ri);

                if (rowStartIndex < 0)
                {
                    rowStartIndex = ~rowStartIndex;

                    if (rowStartIndex > 0)
                    {
                        rowStartIndex--;
                    }
                }

                int delete = (noRows << 20) | noRows;

                for (int i = rowStartIndex; i < rows.Count; i++)
                {
                    ri = rows[i];
                    int fromRowRangeItem = (int)(ri.RowSpan >> 20) + 1;
                    int toRowRangeItem = (int)(ri.RowSpan & 0xFFFFF) + 1;

                    if (fromRowRangeItem >= fromRow)
                    {
                        if (fromRowRangeItem >= fromRow && toRowRangeItem <= fromRow + noRows)
                        {
                            rows.RemoveAt(i--);

                            continue;
                        }
                        else if (fromRowRangeItem >= fromRow + noRows)
                        {
                            if (shiftRow)
                            {
                                toRowRangeItem -= noRows;
                                fromRowRangeItem -= noRows;
                            }
                        }
                        else
                        {
                            if (shiftRow)
                            {
                                fromRowRangeItem = Math.Max(fromRow, fromRowRangeItem - noRows);
                                toRowRangeItem = Math.Max(fromRow, toRowRangeItem - noRows);
                            }
                            else
                            {
                                fromRowRangeItem = Math.Max(fromRowRangeItem, fromRow + noRows + 1);
                                toRowRangeItem = Math.Max(toRowRangeItem, fromRow + noRows + 1);
                            }
                        }
                    }
                    else if (toRowRangeItem >= fromRow)
                    {
                        toRowRangeItem = Math.Max(fromRow, toRowRangeItem - noRows);
                    }

                    ri.RowSpan = ((fromRowRangeItem - 1) << 20) | (toRowRangeItem - 1);
                    rows[i] = ri;
                }
            }
        }
    }

    internal void InsertColumn(int fromCol, int noCols, int fromRow = 1, int toRow = ExcelPackage.MaxRows)
    {
        //Full column
        if (fromRow <= 1 && toRow >= ExcelPackage.MaxRows)
        {
            this.AddFullColumn(fromCol, noCols);
        }
        else
        {
            this.InsertPartialColumn(fromCol, noCols, fromRow, toRow);
        }

        if (this._extendValuesToInsertedColumn)
        {
            this.ExtendValues(fromCol - 1, fromCol + noCols, fromRow, toRow);
        }
    }

    private void ExtendValues(int fromCol, int toCol, int fromRow, int toRow)
    {
        if (this._addresses.ContainsKey(fromCol) && this._addresses.ContainsKey(toCol))
        {
            List<RangeItem>? toColumn = this._addresses[toCol];

            foreach (RangeItem item in this._addresses[fromCol])
            {
                int pos = toColumn.BinarySearch(item);

                if (pos < 0)
                {
                    pos = ~pos;
                }

                while (pos >= 0 && pos < toColumn.Count)
                {
                    RangeItem ri = toColumn[pos];
                    int fr = (int)(ri.RowSpan >> 20) + 1;
                    int tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                    if (tr < fromRow || fr > toRow)
                    {
                        break;
                    }

                    GetIntersect(item, toColumn[pos], out fr, out tr);

                    if (fr >= 0)
                    {
                        fr = Math.Max(fr, fromRow);
                        tr = Math.Min(tr, toRow);
                        this.Add(fr, fromCol + 1, tr, toCol - 1, item.Value);
                    }

                    pos++;
                }
            }
        }
    }

    private static void GetIntersect(RangeItem itemFirst, RangeItem itemLast, out int fr, out int tr)
    {
        if (itemFirst.Value.Equals(itemLast.Value) == false)
        {
            fr = -1;
            tr = -1;

            return;
        }

        int fr1 = (int)(itemFirst.RowSpan >> 20) + 1;
        int tr1 = (int)(itemFirst.RowSpan & 0xFFFFF) + 1;

        int fr2 = (int)(itemLast.RowSpan >> 20) + 1;
        int tr2 = (int)(itemLast.RowSpan & 0xFFFFF) + 1;

        fr = Math.Max(fr1, fr2);
        tr = Math.Min(tr1, tr2);
    }

    internal void DeleteColumn(int fromCol, int noCols, int fromRow = 1, int toRow = ExcelPackage.MaxRows)
    {
        //Full column
        if (fromRow <= 1 && toRow >= ExcelPackage.MaxRows)
        {
            this.DeleteFullColumn(fromCol, noCols);
        }
        else
        {
            this.DeletePartialColumn(fromCol, noCols, fromRow, toRow);
        }
    }

    private void DeletePartialColumn(int fromCol, int noCols, int fromRow, int toRow)
    {
        IOrderedEnumerable<int>? cols = this.GetColumnKeys().OrderBy(x => x);
        int toCol = fromCol + noCols - 1;

        foreach (int colNo in cols)
        {
            if (colNo >= fromCol)
            {
                if (colNo > toCol)
                {
                    this.MoveDataToColumn(colNo, noCols, fromRow, toRow);
                }

                this.DeleteRowsInColumn(colNo, fromRow, toRow);
            }
        }
    }

    private void MoveDataToColumn(int colNo, int noCols, int fromRow, int toRow)
    {
        int destColNo = colNo - noCols;

        if (this._addresses.TryGetValue(colNo, out List<RangeItem> sourceCol))
        {
            for (int i = 0; i < sourceCol.Count; i++)
            {
                RangeItem ri = sourceCol[i];
                int fr = (int)(ri.RowSpan >> 20) + 1;
                int tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                if (fr >= fromRow && tr <= toRow)
                {
                    this.Add(fr, destColNo, tr, destColNo, ri.Value);
                }
                else if (fr <= fromRow && tr >= fromRow)
                {
                    this.Add(fromRow, destColNo, Math.Min(toRow, tr), destColNo, ri.Value);
                }
            }
        }
    }

    private void DeleteRowsInColumn(int colNo, int fromRow, int toRow)
    {
        List<RangeItem>? deleteCol = this._addresses[colNo];

        for (int i = 0; i < deleteCol.Count; i++)
        {
            RangeItem ri = deleteCol[i];
            int fr = (int)(ri.RowSpan >> 20) + 1;
            int tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

            if (fr >= fromRow && tr <= toRow)
            {
                int rows = tr - fr + 1;
                this.DeleteRow(fromRow, tr, colNo, colNo);
                i--;
            }
            else if (tr >= fromRow)
            {
                int ntr = fromRow - 1;
                ri.RowSpan = ri.RowSpan = ((fr - 1) << 20) | (ntr - 1);
                deleteCol[i] = ri;

                if (toRow < tr)
                {
                    this.Add(toRow + 1, colNo, tr, colNo, ri.Value);
                    i++;
                }
            }
        }
    }

    private void InsertPartialColumn(int fromCol, int noCols, int fromRow, int toRow)
    {
        List<int>? cols = this.GetColumnKeys();

        foreach (int colNo in cols.OrderByDescending(x => x))
        {
            if (colNo >= fromCol)
            {
                List<RangeItem>? sourceCol = this._addresses[colNo];

                for (int i = 0; i < sourceCol.Count; i++)
                {
                    RangeItem ri = sourceCol[i];
                    int fr = (int)(ri.RowSpan >> 20) + 1;
                    int tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                    if (fr >= fromRow && tr <= toRow)
                    {
                        int rows = tr - fr + 1;
                        this.DeleteRow(fromRow, tr, colNo, colNo);
                        this.Add(fr, colNo + noCols, tr, colNo + noCols, ri.Value);
                        i--;
                    }
                    else if (tr >= fromRow)
                    {
                        int ntr = fromRow - 1;
                        ri.RowSpan = ri.RowSpan = ((fr - 1) << 20) | (ntr - 1);
                        sourceCol[i] = ri;

                        this.Add(fromRow, colNo + noCols, toRow, colNo + noCols, ri.Value);

                        if (toRow < tr)
                        {
                            this.Add(toRow + 1, colNo, tr, colNo, ri.Value);
                            i++;
                        }
                    }
                }
            }
        }
    }

    private void DeleteFullColumn(int fromCol, int noCols)
    {
        List<int>? cols = this.GetColumnKeys();

        foreach (int key in cols.OrderBy(x => x))
        {
            if (key >= fromCol)
            {
                if (key < fromCol + noCols)
                {
                    this._addresses.Remove(key);
                }
                else
                {
                    List<RangeItem>? col = this._addresses[key];
                    this._addresses.Remove(key);
                    this._addresses.Add(key - noCols, col);
                }
            }
        }
    }

    private void AddFullColumn(int fromCol, int noCols)
    {
        List<int>? cols = this.GetColumnKeys();

        foreach (int key in cols.OrderByDescending(x => x))
        {
            if (key >= fromCol)
            {
                List<RangeItem>? col = this._addresses[key];
                this._addresses.Remove(key);
                this._addresses.Add(key + noCols, col);
            }
        }
    }

    private List<int> GetColumnKeys()
    {
        List<int>? cols = new List<int>();

        foreach (int key in this._addresses.Keys)
        {
            cols.Add(key);
        }

        return cols;
    }

    private static bool ExistsInSpan(int fromRow, int toRow, long r)
    {
        int fr = (int)(r >> 20) + 1;
        int tr = (int)(r & 0xFFFFF) + 1;

        return fr <= toRow && tr >= fromRow;
    }

    private void AddRowSpan(int col, int fromRow, int toRow, T value)
    {
        long rowSpan = ((long)(fromRow - 1) << 20) | (long)(toRow - 1);

        if (this._addresses.TryGetValue(col, out List<RangeItem> rows) == false)
        {
            rows = new List<RangeItem>(64);
            this._addresses.Add(col, rows);
        }

        if (rows.Count == 0)
        {
            rows.Add(new RangeItem(rowSpan, value));

            return;
        }

        RangeItem ri = new RangeItem(rowSpan, value);
        int ix = rows.BinarySearch(ri);

        if (ix < 0)
        {
            ix = ~ix;

            if (ix < rows.Count)
            {
                rows.Insert(ix, ri);
            }
            else
            {
                rows.Add(ri);
            }
        }
    }

    private void MergeRowSpan(int col, int fromRow, int toRow, T value)
    {
        long rowSpan = ((long)(fromRow - 1) << 20) | (long)(toRow - 1);

        if (this._addresses.TryGetValue(col, out List<RangeItem> rows) == false)
        {
            rows = new List<RangeItem>(64);
            this._addresses.Add(col, rows);
        }

        if (rows.Count == 0)
        {
            rows.Add(new RangeItem(rowSpan, value));

            return;
        }

        RangeItem ri = new RangeItem(rowSpan, value);
        int ix = rows.BinarySearch(ri);

        if (ix < 0)
        {
            ix = ~ix;

            if (ix > 0)
            {
                ix--;
            }

            if (ix < rows.Count)
            {
                int tr = -1;

                while (rows.Count > ix)
                {
                    RangeItem rs = rows[ix];
                    int fr = (int)(rs.RowSpan >> 20) + 1;
                    tr = (int)(rs.RowSpan & 0xFFFFF) + 1;

                    if (fr <= fromRow && tr >= toRow)
                    {
                        break; //Inside, exit
                    }

                    if (fr > toRow)
                    {
                        rows.Insert(ix, new RangeItem(rowSpan, value));
                        ix++;

                        break;
                    }
                    else if (fromRow < fr)
                    {
                        rowSpan = ((long)(fromRow - 1) << 20) | (long)(fr - 2);
                        rows.Insert(ix, new RangeItem(rowSpan, value));
                        ix++;
                        fromRow = tr + 1;
                    }

                    ix++;
                }

                if (tr < toRow)
                {
                    tr = tr > fromRow - 1 ? tr : fromRow - 1;
                    rowSpan = ((long)tr << 20) | (long)(toRow - 1);
                    rows.Insert(ix, new RangeItem(rowSpan, value));
                }
            }
            else
            {
                rows.Add(ri);
            }
        }
    }
}