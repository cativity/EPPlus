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
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A pivot table data field
/// </summary>
public class ExcelPivotTableDataField : XmlHelper
{
    internal ExcelPivotTableDataField(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field)
        : base(ns, topNode)
    {
        if (topNode.Attributes.Count == 0)
        {
            this.Index = field.Index;
            this.BaseField = 0;
            this.BaseItem = 0;
        }

        this.Field = field;
    }

    /// <summary>
    /// The field
    /// </summary>
    public ExcelPivotTableField Field { get; internal set; }

    /// <summary>
    /// The index of the datafield
    /// </summary>
    public int Index
    {
        get { return this.GetXmlNodeInt("@fld"); }
        internal set { this.SetXmlNodeString("@fld", value.ToString()); }
    }

    /// <summary>
    /// The name of the datafield
    /// </summary>
    public string Name
    {
        get { return this.GetXmlNodeString("@name"); }
        set
        {
            if (this.Field._pivotTable.DataFields.ExistsDfName(value, this))
            {
                throw new InvalidOperationException("Duplicate datafield name");
            }

            this.SetXmlNodeString("@name", value);
        }
    }

    /// <summary>
    /// Field index. Reference to the field collection
    /// </summary>
    public int BaseField
    {
        get { return this.GetXmlNodeInt("@baseField"); }
        set { this.SetXmlNodeString("@baseField", value.ToString()); }
    }

    /// <summary>
    /// The index to the base item when the ShowDataAs calculation is in use
    /// </summary>
    public int BaseItem
    {
        get { return this.GetXmlNodeInt("@baseItem"); }
        set { this.SetXmlNodeString("@baseItem", value.ToString()); }
    }

    /// <summary>
    /// Number format id. 
    /// </summary>
    internal int NumFmtId
    {
        get { return this.GetXmlNodeInt("@numFmtId"); }
        set { this.SetXmlNodeString("@numFmtId", value.ToString()); }
    }

    /// <summary>
    /// The number format for the data field
    /// </summary>
    public string Format
    {
        get
        {
            foreach (ExcelNumberFormatXml? nf in this.Field._pivotTable.WorkSheet.Workbook.Styles.NumberFormats)
            {
                if (nf.NumFmtId == this.NumFmtId)
                {
                    return nf.Format;
                }
            }

            return this.Field._pivotTable.WorkSheet.Workbook.Styles.NumberFormats[0].Format;
        }
        set
        {
            ExcelStyles? styles = this.Field._pivotTable.WorkSheet.Workbook.Styles;

            ExcelNumberFormatXml nf = null;

            if (!styles.NumberFormats.FindById(value, ref nf))
            {
                nf = new ExcelNumberFormatXml(this.NameSpaceManager) { Format = value, NumFmtId = styles.NumberFormats.NextId++ };
                styles.NumberFormats.Add(value, nf);
            }

            this.NumFmtId = nf.NumFmtId;
        }
    }

    /// <summary>
    /// Type of aggregate function
    /// </summary>
    public DataFieldFunctions Function
    {
        get
        {
            string s = this.GetXmlNodeString("@subtotal");

            if (s == "")
            {
                return DataFieldFunctions.None;
            }
            else
            {
                return (DataFieldFunctions)Enum.Parse(typeof(DataFieldFunctions), s, true);
            }
        }
        set
        {
            string v;

            switch (value)
            {
                case DataFieldFunctions.None:
                    this.DeleteNode("@subtotal");

                    return;

                case DataFieldFunctions.CountNums:
                    v = "countNums";

                    break;

                case DataFieldFunctions.StdDev:
                    v = "stdDev";

                    break;

                case DataFieldFunctions.StdDevP:
                    v = "stdDevP";

                    break;

                default:
                    v = value.ToString().ToLower(CultureInfo.InvariantCulture);

                    break;
            }

            this.SetXmlNodeString("@subtotal", v);
        }
    }

    ExcelPivotTableDataFieldShowDataAs _showDataAs = null;

    /// <summary>
    /// Represents a pivot fields Show As properties.
    /// </summary>
    public ExcelPivotTableDataFieldShowDataAs ShowDataAs
    {
        get { return this._showDataAs ??= new ExcelPivotTableDataFieldShowDataAs(this); }
    }

    internal eShowDataAs ShowDataAsInternal
    {
        get
        {
            string s = this.GetXmlNodeString("@showDataAs");

            if (s == "")
            {
                s = this.GetXmlNodeString("d:extLst/d:ext[@uri='{E15A36E0-9728-4e99-A89B-3F7291B0FE68}']/x14:dataField/@pivotShowAs");

                if (s == "")
                {
                    return eShowDataAs.Normal;
                }
            }

            return s.ToShowDataAs();
        }
        set
        {
            if (value == eShowDataAs.Normal)
            {
                this.DeleteNode("@showDataAs");
            }
            else
            {
                if (IsShowDataAsExtLst(value))
                {
                    this.DeleteNode("@showDataAs");
                    XmlNode? extNode = this.GetOrCreateExtLstSubNode("{E15A36E0-9728-4e99-A89B-3F7291B0FE68}", "x14");
                    XmlHelper? extNodeHelper = XmlHelperFactory.Create(this.NameSpaceManager, extNode);

                    extNodeHelper.SetXmlNodeString("x14:dataField/@pivotShowAs", value.FromShowDataAs());
                }
                else
                {
                    this.DeleteNode("d:extLst/d:ext[@url='{E15A36E0-9728-4e99-A89B-3F7291B0FE68}']");
                    this.SetXmlNodeString("@showDataAs", value.FromShowDataAs());
                }
            }
        }
    }

    private static bool IsShowDataAsExtLst(eShowDataAs value)
    {
        return value == eShowDataAs.PercentOfParent
               || value == eShowDataAs.PercentOfParentColumn
               || value == eShowDataAs.PercentOfParentRow
               || value == eShowDataAs.RankAscending
               || value == eShowDataAs.RankDescending
               || value == eShowDataAs.PercentOfRunningTotal;
    }
}