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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// Defines a pivot table area of selection used for different purposes.
/// </summary>
public class ExcelPivotArea : XmlHelper
{
    ExcelPivotTable _pt;
    internal ExcelPivotArea(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt) : 
        base(nsm, topNode)
    {
        this._pt = pt;
    }
    /// <summary>
    /// The field referenced. -2 means refers to values.
    /// </summary>
    public int? FieldIndex
    { 
        get
        {
            return this.GetXmlNodeInt("@field");
        }
        set
        {
            if(value != null && !(value >= -2 && value< this._pt.Fields.Count))
            {
                throw new InvalidOperationException("Field index out out of range. Field index must be -2 (values) or within the index of the PivotTable's Fields collection");
            }

            this.SetXmlNodeInt("@field", value);
        }
    }
    /// <summary>
    /// Position of the field within the axis to which this rule applies. 
    /// </summary>
    public int? FieldPosition 
    {
        get
        {
            return this.GetXmlNodeIntNull("@fieldPosition");
        }
        set
        {
            if (value != null &&  (value < 0 || value > 255))
            {
                throw new InvalidOperationException("FieldPosition cant be negative and may not exceed 255");
            }

            this.SetXmlNodeInt("@fieldPosition", value);
        }
    }
    /// <summary>
    /// If the pivot area referes to the "Σ Values" field in the column or row fields.
    /// </summary>
    public bool IsValuesField
    {
        get
        {
            return this.FieldIndex == -2;
        }
        set
        {
            this.FieldIndex = -2;
        }
    }
    /// <summary>
    /// The pivot area type that affecting the selection.
    /// </summary>
    public ePivotAreaType PivotAreaType
    {
        get
        {
            return this.GetXmlNodeString("@type").ToPivotAreaType();
        }
        internal set
        {
            if(value==ePivotAreaType.Normal)
            {
                ((XmlElement)this.TopNode).RemoveAttribute("@type");
            }
            else
            {
                this.SetXmlNodeString("@type", value.ToPivotAreaTypeString());
            }
        }
    }
    /// <summary>
    /// The region of the PivotTable affected.
    /// </summary>
    public ePivotTableAxis Axis 
    { 
        get
        {
            return this.GetXmlNodeString("@axis").ToPivotTableAxis();
        }
        set
        {
            this.SetXmlNodeString("@axis", value.ToPivotTableAxisString(), true);
        }
    }

    /// <summary>
    /// If the data values in the data area are included. Setting this property to true will set <see cref="LabelOnly"/> to false.
    /// <seealso cref="LabelOnly"/>
    /// </summary>
    public bool DataOnly 
    { 
        get
        {
            return this.GetXmlNodeBool("@dataOnly", true);
        }
        set
        {
            if (value && (this.PivotAreaType == ePivotAreaType.Data || this.PivotAreaType == ePivotAreaType.Normal || this.PivotAreaType == ePivotAreaType.Origin || this.PivotAreaType == ePivotAreaType.TopEnd))
            {
                throw new InvalidOperationException("Can't set LabelOnly to True for the PivotAreaType");
            }
            if (value && this.LabelOnly)
            {
                this.LabelOnly = false;
            }

            this.SetXmlNodeBool("@dataOnly", value, true);
        }
    }
    /// <summary>
    /// If the item labels are included. Setting this property to true will set <see cref="DataOnly"/> to false.
    /// <seealso cref="DataOnly"/>
    /// </summary>
    public bool LabelOnly
    {
        get
        {
            return this.GetXmlNodeBool("@labelOnly");
        }
        set
        {
            if(value && this.DataOnly)
            {
                this.DataOnly = false;
            }

            this.SetXmlNodeBool("@labelOnly", value);
        }
    }
    /// <summary>
    /// If the row grand total is included
    /// </summary>
    public bool GrandRow
    {
        get
        {
            return this.GetXmlNodeBool("@grandRow");
        }
        set
        {
            this.SetXmlNodeBool("@grandRow", value);
        }
    }
    /// <summary>
    /// If the column grand total is included
    /// </summary>
    public bool GrandColumn
    {
        get
        {
            return this.GetXmlNodeBool("@grandCol");
        }
        set
        {
            this.SetXmlNodeBool("@grandCol", value);
        }
    }
    /// <summary>
    /// If any indexes refers to fields or items in the pivot cache and not the view.
    /// </summary>
    public bool CacheIndex
    {
        get
        {
            return this.GetXmlNodeBool("@cacheIndex", true);
        }
        set
        {
            this.SetXmlNodeBool("@cacheIndex", value, true);
        }
    }
    /// <summary>
    /// Indicating whether the pivot table area refers to an area that is in outline mode.
    /// </summary>
    public bool Outline
    {
        get
        {
            return this.GetXmlNodeBool("@outline", true);
        }
        set
        {
            this.SetXmlNodeBool("@outline", value, true);
        }
    }
    /// <summary>
    /// A address in A1C1 format that specifies a subset of the selection area. Points are relative to the top left of the selection area.
    /// The first cell is referenced as A1. For example, B1:C1 reference the second and third column of the first row of the pivot area.
    /// </summary>
    public string Offset
    {
        get
        {
            return this.GetXmlNodeString("@offset");
        }
        internal set
        {
            this.SetXmlNodeString("@offset", value, true);
        }
    }
    /// <summary>
    /// If collapsed levels/dimensions are considered subtotals
    /// </summary>
    public bool CollapsedLevelsAreSubtotals 
    {
        get
        {
            return this.GetXmlNodeBool("@collapsedLevelsAreSubtotals");
        }
        set
        {
            this.SetXmlNodeBool("@collapsedLevelsAreSubtotals", value, false);
        }
    }
}