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

using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.DataValidation.Formulas;

internal class ExcelDataValidationFormulaList : ExcelDataValidationFormula, IExcelDataValidationFormulaList
{
    #region class DataValidationList

    private class DataValidationList : IList<string>, ICollection
    {
        private IList<string> _items = new List<string>();
        private EventHandler<EventArgs> _listChanged;

        public event EventHandler<EventArgs> ListChanged
        {
            add { this._listChanged += value; }
            remove { this._listChanged -= value; }
        }

        private void OnListChanged()
        {
            if (this._listChanged != null)
            {
                this._listChanged(this, EventArgs.Empty);
            }
        }

        #region IList members

        int IList<string>.IndexOf(string item)
        {
            return this._items.IndexOf(item);
        }

        void IList<string>.Insert(int index, string item)
        {
            this._items.Insert(index, item);
            this.OnListChanged();
        }

        void IList<string>.RemoveAt(int index)
        {
            this._items.RemoveAt(index);
            this.OnListChanged();
        }

        string IList<string>.this[int index]
        {
            get { return this._items[index]; }
            set
            {
                this._items[index] = value;
                this.OnListChanged();
            }
        }

        void ICollection<string>.Add(string item)
        {
            this._items.Add(item);
            this.OnListChanged();
        }

        void ICollection<string>.Clear()
        {
            this._items.Clear();
            this.OnListChanged();
        }

        bool ICollection<string>.Contains(string item)
        {
            return this._items.Contains(item);
        }

        void ICollection<string>.CopyTo(string[] array, int arrayIndex)
        {
            this._items.CopyTo(array, arrayIndex);
        }

        int ICollection<string>.Count
        {
            get { return this._items.Count; }
        }

        bool ICollection<string>.IsReadOnly
        {
            get { return false; }
        }

        bool ICollection<string>.Remove(string item)
        {
            bool retVal = this._items.Remove(item);
            this.OnListChanged();

            return retVal;
        }

        IEnumerator<string> IEnumerable<string>.GetEnumerator()
        {
            return this._items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this._items.GetEnumerator();
        }

        #endregion

        public void CopyTo(Array array, int index)
        {
            this._items.CopyTo((string[])array, index);
        }

        int ICollection.Count
        {
            get { return this._items.Count; }
        }

        public bool IsSynchronized
        {
            get { return ((ICollection)this._items).IsSynchronized; }
        }

        public object SyncRoot
        {
            get { return ((ICollection)this._items).SyncRoot; }
        }
    }

    #endregion

    public ExcelDataValidationFormulaList(string formula, string uid, string sheetName, Action<OnFormulaChangedEventArgs> extListHandler)
        : base(uid, sheetName, extListHandler)
    {
        DataValidationList? values = new DataValidationList();
        values.ListChanged += new EventHandler<EventArgs>(this.values_ListChanged);
        this.Values = values;
        this._inputFormula = formula;
        this.SetInitialValues();
    }

    private string _inputFormula;

    private void SetInitialValues()
    {
        string? @value = this._inputFormula;

        if (!string.IsNullOrEmpty(@value))
        {
            if (@value.StartsWith("\"", StringComparison.OrdinalIgnoreCase) && @value.EndsWith("\"", StringComparison.OrdinalIgnoreCase))
            {
                @value = @value.TrimStart('"').TrimEnd('"');
                string[]? items = @value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string? item in items)
                {
                    this.Values.Add(item);
                }
            }
            else
            {
                this.ExcelFormula = @value;
            }
        }
    }

    void values_ListChanged(object sender, EventArgs e)
    {
        if (this.Values.Count > 0)
        {
            this.State = FormulaState.Value;
        }

        string? valuesAsString = this.GetValueAsString();

        // Excel supports max 255 characters in this field.
        if (valuesAsString.Length > 255)
        {
            throw new InvalidOperationException("The total length of a DataValidation list cannot exceed 255 characters");
        }
    }

    public IList<string> Values { get; private set; }

    protected override string GetValueAsString()
    {
        StringBuilder? sb = new StringBuilder();

        foreach (string? val in this.Values)
        {
            if (sb.Length == 0)
            {
                _ = sb.Append("\"");
                _ = sb.Append(val);
            }
            else
            {
                _ = sb.AppendFormat(",{0}", val);
            }
        }

        _ = sb.Append("\"");

        return sb.ToString();
    }

    internal override void ResetValue()
    {
        this.Values.Clear();
    }
}