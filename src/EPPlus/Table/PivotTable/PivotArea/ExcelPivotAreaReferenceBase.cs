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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// Base class for pivot area references
/// </summary>
public abstract class ExcelPivotAreaReferenceBase : XmlHelper
{
    internal ExcelPivotTable _pt;
    internal ExcelPivotAreaReferenceBase(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt) : base(nsm, topNode)
    {
        this._pt = pt;
    }
    internal int FieldIndex
    { 
        get
        {
            long v= this.GetXmlNodeLong("@field");
            if(v > int.MaxValue)
            {
                return -2;
            }
            else
            {
                return (int)v;
            }
        }
        set
        {
            if(value<0)
            {
                this.SetXmlNodeLong("@field", 4294967294);
            }
            else
            {
                this.SetXmlNodeInt("@field", value);
            }
        }
    }
    /// <summary>
    /// If this field has selection. This property is used when the pivot table is in outline view. It is also used when both header and data cells have selection.
    /// </summary>
    public bool Selected 
    {
        get
        {
            return this.GetXmlNodeBool("@selected", true);
        }
        set
        {
            this.SetXmlNodeBool("@selected", value);
        }
    }
    /// <summary>
    /// If the item is referred to by a relative reference rather than an absolute reference.
    /// </summary>
    internal bool Relative 
    { 
        get
        {
            return this.GetXmlNodeBool("@relative");
        }
        set
        {
            this.SetXmlNodeBool("@relative", value);
        }
    }
    /// <summary>
    /// Whether the item is referred to by position rather than item index.
    /// </summary>
    internal bool ByPosition 
    {
        get
        {
            return this.GetXmlNodeBool("@byPosition");
        }
        set
        {
            this.SetXmlNodeBool("@byPosition", value);
        }
    }
    internal abstract void UpdateXml();
    /// <summary>
    /// If the default subtotal is included in the filter.
    /// </summary>
    public bool DefaultSubtotal 
    { 
        get
        {
            return this.GetXmlNodeBool("@defaultSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@defaultSubtotal", value);
        }
    }
    /// <summary>
    /// If the Average aggregation function is included in the filter.
    /// </summary>
    public bool AvgSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@avgSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@avgSubtotal", value);
        }
    }
    /// <summary>
    /// If the Count aggregation function is included in the filter.
    /// </summary>
    public bool CountSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@countSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@countSubtotal", value);
        }
    }
    /// <summary>
    /// If the CountA aggregation function is included in the filter.
    /// </summary>
    public bool CountASubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@countASubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@countASubtotal", value);
        }
    }
    /// <summary>
    /// If the Maximum aggregation function is included in the filter.
    /// </summary>
    public bool MaxSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@maxSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@maxSubtotal", value);
        }
    }
    /// <summary>
    /// If the Minimum aggregation function is included in the filter.
    /// </summary>
    public bool MinSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@minSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@minSubtotal", value);
        }
    }
    /// <summary>
    /// If the Product aggregation function is included in the filter.
    /// </summary>
    public bool ProductSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@productSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@productSubtotal", value);
        }
    }
    /// <summary>
    /// If the population standard deviation aggregation function is included in the filter.
    /// </summary>
    public bool StdDevPSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@StdDevPSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@StdDevPSubtotal", value);
        }
    }
    /// <summary>
    /// If the standard deviation aggregation function is included in the filter.
    /// </summary>
    public bool StdDevSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@StdDevSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@StdDevSubtotal", value);
        }
    }
    /// <summary>
    /// If the sum aggregation function is included in the filter.
    /// </summary>
    public bool SumSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@sumSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@sumSubtotal", value);
        }
    }
    /// <summary>
    /// If the population variance aggregation function is included in the filter.
    /// </summary>
    public bool VarPSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@varPSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@varPSubtotal", value);
        }
    }
    /// <summary>
    /// If the variance aggregation function is included in the filter.
    /// </summary>
    public bool VarSubtotal
    {
        get
        {
            return this.GetXmlNodeBool("@varSubtotal");
        }
        set
        {
            this.SetXmlNodeBool("@varSubtotal", value);
        }
    }
    internal void SetFunction(DataFieldFunctions function)
    {
        switch(function)
        {
            case DataFieldFunctions.Average:
                this.AvgSubtotal = true;
                break;
            case DataFieldFunctions.Count:
                this.CountSubtotal = true;
                break;
            case DataFieldFunctions.CountNums:
                this.CountASubtotal = true;
                break;
            case DataFieldFunctions.Max:
                this.MaxSubtotal = true;
                break;
            case DataFieldFunctions.Min:
                this.MinSubtotal = true;
                break;
            case DataFieldFunctions.Product:
                this.ProductSubtotal = true;
                break;
            case DataFieldFunctions.StdDevP:
                this.StdDevPSubtotal = true;
                break;
            case DataFieldFunctions.StdDev:
                this.StdDevSubtotal = true;
                break;
            case DataFieldFunctions.Sum:
                this.SumSubtotal = true;
                break;
            case DataFieldFunctions.VarP:
                this.VarPSubtotal = true;
                break;
            case DataFieldFunctions.Var:
                this.VarSubtotal = true;
                break;
            default:
                this.DefaultSubtotal = true;
                break;
        }
    }
}