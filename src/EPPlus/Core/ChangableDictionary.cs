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
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Core;

internal class ChangeableDictionary<T> : IEnumerable<T>
{
    internal int[][] _index;
    internal List<T> _items;
    internal int _count;
    int _defaultSize;
    internal ChangeableDictionary(int size = 8)
    {
        this._defaultSize = size;
        this.Clear();
    }

    internal T this[int key]
    {
        get
        {
            int pos = Array.BinarySearch(this._index[0], 0, this._count, key);
            if(pos>=0)
            {
                return this._items[this._index[1][pos]];
            }
            else
            {
                return default(T);
            }
        }
    }

    internal void InsertAndShift(int fromPosition, int add)
    {
        int pos = Array.BinarySearch(this._index[0], 0, this._count, fromPosition);
        if(pos<0)
        {
            pos = ~pos;
        }
        Array.Copy(this._index[0], pos, this._index[0], pos + 1, this._count - pos);
        Array.Copy(this._index[1], pos, this._index[1], pos + 1, this._count - pos);
        this._count++;
        for (int i=pos;i< this.Count;i++)
        {
            this._index[0][i] += add;
        }
    }
        
    internal int Count { get { return this._count; } }

    public void Add(int key, T value)
    {
        int pos = Array.BinarySearch(this._index[0], 0, this._count, key);
        if (pos >= 0)
        {
            throw new ArgumentException("Key already exists");
        }
        pos = ~pos;
        if (pos >= this._index[0].Length - 1)
        {
            Array.Resize(ref this._index[0], this._index[0].Length << 1);
            Array.Resize(ref this._index[1], this._index[1].Length << 1);
        }
        if (pos < this.Count)
        {
            Array.Copy(this._index[0], pos, this._index[0], pos + 1, this._index[0].Length - pos - 1);
            Array.Copy(this._index[1], pos, this._index[1], pos + 1, this._index[1].Length - pos - 1);
        }

        this._count++;
        this._index[0][pos] = key;
        this._index[1][pos] = this._items.Count;
        this._items.Add(value);
    }

    internal void Move(int fromPosition, int toPosition, bool before)
    {
        if (this.Count <= 1 || fromPosition == toPosition)
        {
            return;
        }

        int listItem = this._index[1][fromPosition];
        int insertPos = before ? toPosition : toPosition + 1;

        if(insertPos>fromPosition)
        {
            this.InsertAndShift(insertPos, 1);
            this.RemoveAndShift(fromPosition, false);
            insertPos--;
        }
        else
        {
            this.RemoveAndShift(fromPosition, false);
            this.InsertAndShift(insertPos, 1);
        }

        this._index[0][insertPos] = insertPos;
        this._index[1][insertPos] = listItem;
    }

    public void Clear()
    {
        this._index = new int[2][];
        this._index[0] = new int[this._defaultSize];
        this._index[1] = new int[this._defaultSize];
        this._items = new List<T>();
    }

    public bool ContainsKey(int key)
    {
        return Array.BinarySearch(this._index[0], 0, this._count, key) >= 0;
    }
    
    public IEnumerator<T> GetEnumerator()
    {
        return new ChangeableDictionaryEnumerator<T>(this);
    }

    public bool RemoveAndShift(int key)
    {
        return this.RemoveAndShift(key, true);
    }

    private bool RemoveAndShift(int key, bool dispose)
    {
        int pos = Array.BinarySearch(this._index[0], 0, this._count, key);
        if (pos >= 0)
        {
            if (dispose)
            {
                (this._items[this._index[1][pos]] as IDisposable)?.Dispose();
                this._items[this._index[1][pos]] = default(T);
            }

            if (pos < this.Count)
            {
                Array.Copy(this._index[0], pos + 1, this._index[0], pos, this.Count - pos - 1);
                Array.Copy(this._index[1], pos + 1, this._index[1], pos, this.Count - pos - 1);
            }

            this._count--;
            for (int i = pos; i < this._count; i++)
            {
                this._index[0][i]--;
            }
            return true;
        }
        return false;
    }

    public bool TryGetValue(int key, out T value)
    {
        int pos = Array.BinarySearch(this._index[0], 0, this._count, key);
        if (pos >= 0)
        {
            value = this._items[pos];
            return true;
        }
        else
        {
            value = default(T);
            return false;
        }
    }
    IEnumerator IEnumerable.GetEnumerator()
    {
        return new ChangeableDictionaryEnumerator<T>(this);
    }
}
internal class ChangeableDictionaryEnumerator<T> : IEnumerator<T>
{
    int _index=-1;
    ChangeableDictionary<T> _ts;
    public ChangeableDictionaryEnumerator(ChangeableDictionary<T> ts)
    {
        this._ts = ts;
    }
    public T Current
    {
        get
        {
            if (this._index >= this._ts._count)
            {
                return default(T);
            }
            else
            {
                return this._ts._items[this._ts._index[1][this._index]];
            }
        }
    }

    object IEnumerator.Current => this.Current;

    public void Dispose()
    {
        this._ts = null;
    }

    public bool MoveNext()
    {
        this._index++;
        if (this._ts.Count == this._index)
        {
            return false;
        }
        return true;
    }

    public void Reset()
    {
        throw new NotImplementedException();
    }
}