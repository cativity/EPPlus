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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using System.Linq;
using OfficeOpenXml.Core;
using System.Collections;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A reference to a field in a pivot area 
/// </summary>
public class ExcelPivotAreaReference : ExcelPivotAreaReferenceBase
{
    internal ExcelPivotAreaReference(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt, int fieldIndex = -1)
        : base(nsm, topNode, pt)
    {
        this.Items = new ExcelPivotAreaReferenceItems(this);

        if (fieldIndex != -1)
        {
            this.FieldIndex = fieldIndex;
        }

        if (this.FieldIndex >= 0)
        {
            foreach (XmlNode n in topNode.ChildNodes)
            {
                if (n.LocalName == "x")
                {
                    int ix = int.Parse(n.Attributes["v"].Value);

                    if (ix < this.Field.Items.Count)
                    {
                        this.Items.Add(new PivotItemReference() { Index = ix, Value = this.Field.Items[ix].Value });
                    }
                }
            }
        }
    }

    /// <summary>
    /// The pivot table field referenced
    /// </summary>
    public ExcelPivotTableField Field
    {
        get
        {
            if (this.FieldIndex >= 0)
            {
                return this._pt.Fields[this.FieldIndex];
            }

            return null;
        }
    }

    /// <summary>
    /// References to the pivot table cache or within the table.
    /// </summary>
    public ExcelPivotAreaReferenceItems Items { get; }

    internal override void UpdateXml()
    {
        //Remove reference, so they can be re-written 
        if (this.TopNode.LocalName == "reference")
        {
            while (this.TopNode.ChildNodes.Count > 0)
            {
                _ = this.TopNode.RemoveChild(this.TopNode.ChildNodes[0]);
            }
        }

        if (this.FieldIndex >= 0 && this.FieldIndex < this._pt.Fields.Count)
        {
            ExcelPivotTableFieldItemsCollection? items = this.Field.Items;

            foreach (PivotItemReference r in this.Items)
            {
                if (r.Index >= 0 && r.Index < items.Count && r.Value.Equals(items[r.Index]))
                {
                    XmlElement? n = (XmlElement)this.CreateNode("d:x", false, true);
                    n.SetAttribute("v", r.Index.ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    int ix = items._list.FindIndex(x => (x.Value != null && x.Value.Equals(r.Value)) || (x.Text != null && x.Text.Equals(r.Value)));

                    if (ix >= 0)
                    {
                        XmlElement? n = (XmlElement)this.CreateNode("d:x", false, true);
                        n.SetAttribute("v", ix.ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
        }
    }
}