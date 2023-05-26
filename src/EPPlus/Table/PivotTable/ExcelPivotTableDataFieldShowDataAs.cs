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
namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Compares the item to the previous or next item.
    /// </summary>
    public enum ePrevNextPivotItem
    {
        /// <summary>
        /// The Previous item
        /// </summary>
        Previous = 1048828,
        /// <summary>
        /// The Next item
        /// </summary>
        Next = 1048829
    }

    /// <summary>
    /// Represents a pivot fields Show As properties.
    /// </summary>
    public class ExcelPivotTableDataFieldShowDataAs
    {
        ExcelPivotTableDataField _dataField;
        internal ExcelPivotTableDataFieldShowDataAs(ExcelPivotTableDataField dataField)
        {
            this._dataField = dataField;
        }
        /// <summary>
        /// Sets the show data as to type Normal. This removes the Show data as setting.
        /// </summary>
        public void SetNormal()
        {
            this._dataField.ShowDataAsInternal = eShowDataAs.Normal;
            this._dataField.BaseField = 0;
            this._dataField.BaseItem = 0;
        }

        /// <summary>
        /// Sets the show data as to type Percent Of Total
        /// </summary>
        public void SetPercentOfTotal()
        {
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentOfTotal;
            this._dataField.BaseField = 0;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Row
        /// </summary>
        public void SetPercentOfRow()
        {
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentOfRow;
            this._dataField.BaseField = 0;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Column
        /// </summary>
        public void SetPercentOfColumn()
        {
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentOfColumn;
            this._dataField.BaseField = 0;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Percent
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The index of the item to use within the <see cref="ExcelPivotTableField.Items"/> collection of the base field</param>
        /// </summary>
        public void SetPercent(ExcelPivotTableField baseField, int baseItem)
        {
            this.Validate(baseField, baseItem);
            this._dataField.ShowDataAsInternal = eShowDataAs.Percent;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = baseItem;
        }
        /// <summary>
        /// Sets the show data as to type Percent
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The previous or next field</param>
        /// </summary>
        public void SetPercent(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.Percent;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = (int)baseItem;
        }

        /// <summary>
        /// Sets the show data as to type Percent Of Parent
        /// <param name="baseField">The base field to use</param>
        /// </summary>
        public void SetPercentParent(ExcelPivotTableField baseField)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentOfParent;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = 0;
        }

        /// <summary>
        /// Sets the show data as to type Index
        /// </summary>
        public void SetIndex()
        {
            this._dataField.ShowDataAsInternal = eShowDataAs.Index;
            this._dataField.BaseField = 0;
            this._dataField.BaseItem = 0;
        }

        /// <summary>
        /// Sets the show data as to type Running Total
        /// </summary>
        /// <param name="baseField">The base field to use</param>
        public void SetRunningTotal(ExcelPivotTableField baseField)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.RunningTotal;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Difference
        /// </summary>
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The index of the item to use within the <see cref="ExcelPivotTableField.Items"/> collection of the base field</param>
        public void SetDifference(ExcelPivotTableField baseField, int baseItem)
        {
            this.Validate(baseField, baseItem);
            this._dataField.ShowDataAsInternal = eShowDataAs.Difference;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = baseItem;
        }
        /// <summary>
        /// Sets the show data as to type Difference
        /// </summary>
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The previous or next field</param>
        public void SetDifference(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.Difference;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = (int)baseItem;
        }

        /// <summary>
        /// Sets the show data as to type Percent Of Total
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The index of the item to use within the <see cref="ExcelPivotTableField.Items"/> collection of the base field</param>
        /// </summary>
        public void SetPercentageDifference(ExcelPivotTableField baseField, int baseItem)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentDifference;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = baseItem;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Total
        /// <param name="baseField">The base field to use</param>
        /// <param name="baseItem">The previous or next field</param>
        /// </summary>
        public void SetPercentageDifference(ExcelPivotTableField baseField, ePrevNextPivotItem baseItem)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentDifference;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = (int)baseItem;
        }

        /// <summary>
        /// Sets the show data as to type Percent Of Parent Row
        /// </summary>
        public void SetPercentParentRow()
        {
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentOfParentRow;
            this._dataField.BaseField = 0;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Parent Column
        /// </summary>
        public void SetPercentParentColumn()
        {
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentOfParentColumn;
            this._dataField.BaseField = 0;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Percent Of Running Total
        /// </summary>
        public void SetPercentOfRunningTotal(ExcelPivotTableField baseField)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.PercentOfRunningTotal;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Rank Ascending
        /// <param name="baseField">The base field to use</param>
        /// </summary>
        public void SetRankAscending(ExcelPivotTableField baseField)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.RankAscending;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// Sets the show data as to type Rank Descending
        /// <param name="baseField">The base field to use</param>
        /// </summary>
        public void SetRankDescending(ExcelPivotTableField baseField)
        {
            this.Validate(baseField);
            this._dataField.ShowDataAsInternal = eShowDataAs.RankDescending;
            this._dataField.BaseField = baseField.Index;
            this._dataField.BaseItem = 0;
        }
        /// <summary>
        /// The value of the "Show Data As" setting
        /// </summary>
        public eShowDataAs Value
        {
            get
            {
                return this._dataField.ShowDataAsInternal;
            }
        }
        private void Validate(ExcelPivotTableField baseField, int? baseItem = null)
        {
            if (baseField._pivotTable != this._dataField.Field._pivotTable)
            {
                throw new ArgumentException("The base field must be from the same pivot table as the data field", nameof(baseField));
            }
            if (baseField == this._dataField.Field)
            {
                throw new ArgumentException("The base field and the data field must not be the same.", nameof(baseField));
            }
            if (baseItem != null)
            {
                if (baseItem<0 || baseItem >= baseField.Items.Count)
                {
                    throw new ArgumentException("Base items must be within an index the fields item collection. Please refresh the Items collection of the field to get the items from source.", nameof(baseField));
                }
            }
        }

    }
}