﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/

using OfficeOpenXml.Core;
using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A collection of pivot area references. A pivot area reference is a reference to a column, row field or a data field
/// </summary>
public class ExcelPivotAreaReferenceCollection : EPPlusReadOnlyList<ExcelPivotAreaReference>
{
    XmlHelper _xmlHelper;
    ExcelPivotTable _pt;

    internal ExcelPivotAreaReferenceCollection(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt)
    {
        this._xmlHelper = XmlHelperFactory.Create(nsm, topNode);
        this._pt = pt;
    }

    /// <summary>
    /// Adds a pivot table field to the collection. The field is usually a column or row field
    /// </summary>
    /// <param name="field">The column or row field</param>
    /// <returns>The pivot area reference</returns>
    public ExcelPivotAreaReference Add(ExcelPivotTableField field) => this.Add(field._pivotTable, field.Index);

    /// <summary>
    /// Adds a pivot table field to the collection. The field is usually a column or row field
    /// </summary>
    /// <param name="pivotTable">The pivot table</param>
    /// <param name="fieldIndex">The index of the pivot table field</param>
    /// <returns></returns>
    public ExcelPivotAreaReference Add(ExcelPivotTable pivotTable, int fieldIndex)
    {
        XmlNode? n = this._xmlHelper.CreateNode("d:references");
        XmlNode? rn = this._xmlHelper.CreateNode(n, "d:reference", true);

        if (pivotTable != this._pt)
        {
            throw new InvalidOperationException("The pivot table field is from another pivot table.");
        }

        ExcelPivotAreaReference? r = new ExcelPivotAreaReference(this._xmlHelper.NameSpaceManager, rn, pivotTable, fieldIndex);
        this._list.Add(r);

        return r;
    }
}