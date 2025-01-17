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
using System.Globalization;
using System.Text;
using System.Xml;
using System.Linq;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml;

/// <summary>
/// A collection of named styles in the workbooks styles.
/// </summary>
/// <typeparam name="T">The type of style</typeparam>
public class ExcelNamedStyleCollection<T> : ExcelStyleCollection<T>
{
    /// <summary>
    /// Indexer for the collection
    /// </summary>
    /// <param name="name">The name of the Style</param>
    /// <returns></returns>
    public T this[string name]
    {
        get
        {
            if (this._dic.ContainsKey(name))
            {
                return this._list[this._dic[name]];
            }

            return default(T);
        }
    }
}

/// <summary>
/// Base collection class for styles.
/// </summary>
/// <typeparam name="T">The style type</typeparam>
public class ExcelStyleCollection<T> : IEnumerable<T>
{
    internal ExcelStyleCollection() => this._setNextIdManual = false;

    bool _setNextIdManual;

    internal ExcelStyleCollection(bool SetNextIdManual) => this._setNextIdManual = SetNextIdManual;

    /// <summary>
    /// The top xml node of the collection
    /// </summary>
    public XmlNode TopNode { get; set; }

    internal List<T> _list = new List<T>();
    internal Dictionary<string, int> _dic = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    internal int NextId;

    #region IEnumerable<T> Members

    /// <summary>
    /// Returns an enumerator that iterates through a collection.
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<T> GetEnumerator() => this._list.GetEnumerator();

    #endregion

    #region IEnumerable Members

    /// <summary>
    /// Returns an enumerator that iterates through a collection.
    /// </summary>
    /// <returns>The enumerator</returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => this._list.GetEnumerator();

    #endregion

    /// <summary>
    /// Indexer for the collection
    /// </summary>
    /// <param name="PositionID">The index of the Style</param>
    /// <returns></returns>
    public T this[int PositionID]
    {
        get
        {
            if (PositionID < 0)
            {
                return default;
            }

            return this._list[PositionID];
        }
    }

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count => this._list.Count;

    internal int Add(string key, T item)
    {
        this._list.Add(item);

        if (!this._dic.ContainsKey(key.ToLower(CultureInfo.InvariantCulture)))
        {
            this._dic.Add(key.ToLower(CultureInfo.InvariantCulture), this._list.Count - 1);
        }

        if (this._setNextIdManual)
        {
            this.NextId++;
        }

        return this._list.Count - 1;
    }

    /// <summary>
    /// Finds the key 
    /// </summary>
    /// <param name="key">the key to be found</param>
    /// <param name="obj">The found object.</param>
    /// <returns>True if found</returns>
    internal bool FindById(string key, ref T obj)
    {
        if (this._dic.ContainsKey(key))
        {
            obj = this._list[this._dic[key.ToLower(CultureInfo.InvariantCulture)]];

            return true;
        }
        else
        {
            return false;
        }
    }

    /// <summary>
    /// Find Index
    /// </summary>
    /// <param name="key"></param>
    /// <returns></returns>
    internal int FindIndexById(string key)
    {
        if (this._dic.ContainsKey(key))
        {
            return this._dic[key];
        }
        else
        {
            return int.MinValue;
        }
    }

    internal int FindIndexByBuildInId(int id)
    {
        for (int i = 0; i < this._list.Count; i++)
        {
            if (this._list[i] is ExcelNamedStyleXml ns)
            {
                if (ns.BuildInId == id)
                {
                    return i;
                }
            }
        }

        return -1;
    }

    internal bool ExistsKey(string key) => this._dic.ContainsKey(key);

    internal void Sort(Comparison<T> c) => this._list.Sort(c);
}