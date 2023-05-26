/*************************************************************************************************
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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using System.Collections;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A collection of data fields used in a pivot area selection
/// </summary>
public class ExcelPivotAreaDataFieldReference : ExcelPivotAreaReferenceBase, IEnumerable<ExcelPivotTableDataField>
{
    List<ExcelPivotTableDataField> _dataFields = new List<ExcelPivotTableDataField>();
    internal ExcelPivotAreaDataFieldReference(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt, int fieldIndex = -1) : base(nsm, topNode, pt)
    {
        if(this.TopNode.LocalName=="reference")
        {
            foreach (XmlNode c in this.TopNode.ChildNodes)
            {
                if (c.LocalName == "x")
                {
                    int ix = int.Parse(c.Attributes["v"].Value);
                    if (ix < pt.DataFields.Count)
                    {
                        this._dataFields.Add(pt.DataFields[ix]);
                    }
                }
            }
        }
    }
    /// <summary>
    /// The indexer
    /// </summary>
    /// <param name="index">The zero-based index of the collection</param>
    /// <returns></returns>
    public ExcelPivotTableDataField this[int index]
    {
        get
        {
            return this._dataFields[index];
        }
    }
    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count 
    { 
        get
        {
            return this._dataFields.Count;
        }
    }
    internal void AddInternal(ExcelPivotTableDataField item)
    {
        this._dataFields.Add(item);
    }
    /// <summary>
    /// Adds the data field at the specific index
    /// </summary>
    /// <param name="index"></param>
    public void Add(int index)
    {
        if (index >= 0 && index < this._pt.DataFields.Count)
        {
            this._dataFields.Add(this._pt.DataFields[index]);
        }
        else
        {
            throw new IndexOutOfRangeException("Index is out of range for referenced data field.");
        }
    }
    /// <summary>
    /// Adds a data field from the pivot table to the pivot area
    /// </summary>
    /// <param name="field"></param>
    public void Add(ExcelPivotTableDataField field)
    {
        if (field == null)
        {
            throw new ArgumentNullException("The pivot table field must not be null.");
        }
        if (field.Field._pivotTable != this._pt)
        {
            throw new ArgumentException("The pivot table field is from another pivot table.");
        }

        this._dataFields.Add(field);
    }

    internal override void UpdateXml()
    {
        //Remove reference, so they can be re-written 
        if (this.TopNode.LocalName == "reference")
        {
            while (this.TopNode.ChildNodes.Count > 0)
            {
                this.TopNode.RemoveChild(this.TopNode.ChildNodes[0]);
            }
        }

        if (this._dataFields.Count==0 && this.FieldIndex>=0)
        {
            if(this.TopNode.LocalName == "reference")
            {
                this.TopNode.ParentNode.ParentNode.RemoveChild(this.TopNode.ParentNode);
            }
            return;
        }
        else
        {
            if (this.TopNode.LocalName == "pivotArea")
            {
                XmlNode? n = this.CreateNode("d:references");
                XmlElement? rn = (XmlElement)this.CreateNode(n, "d:reference", true);
                rn.SetAttribute("field", "4294967294");
                this.TopNode = rn;
            }
        }

        foreach (ExcelPivotTableDataField r in this._dataFields)
        {
            if (r.Field.IsDataField)
            {
                int ix = this._pt.DataFields._list.IndexOf(r);
                if (ix >= 0)
                {
                    XmlElement? n = (XmlElement)this.CreateNode("d:x", false, true);
                    n.SetAttribute("v", ix.ToString(CultureInfo.InvariantCulture));
                }
            }
        }            
    }
    internal void Clear()
    {
        this._dataFields.Clear();
    }
    /// <summary>
    /// Gets the enumerator
    /// </summary>
    /// <returns></returns>
    public IEnumerator<ExcelPivotTableDataField> GetEnumerator()
    {
        return ((IEnumerable<ExcelPivotTableDataField>)this._dataFields).GetEnumerator();
    }

    /// <summary>
    /// Gets the enumerator
    /// </summary>
    /// <returns></returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable)this._dataFields).GetEnumerator();
    }
}