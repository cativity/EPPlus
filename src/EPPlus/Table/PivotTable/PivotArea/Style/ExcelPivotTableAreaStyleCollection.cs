using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A collection of pivot areas used for styling a pivot table.
/// </summary>
public class ExcelPivotTableAreaStyleCollection : EPPlusReadOnlyList<ExcelPivotTableAreaStyle>
{
    ExcelStyles _styles;
    XmlHelper _xmlHelper;
    ExcelPivotTable _pt;

    internal ExcelPivotTableAreaStyleCollection(ExcelPivotTable pt)
    {
        this._pt = pt;
        this._styles = pt.WorkSheet.Workbook.Styles;

        foreach (XmlNode node in pt.GetNodes("d:formats/d:format/d:pivotArea"))
        {
            ExcelPivotTableAreaStyle? s = new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, node, this._pt);
            this._list.Add(s);
        }
    }

    /// <summary>
    /// Adds a pivot area style for labels or data.
    /// </summary>
    /// <returns></returns>
    public ExcelPivotTableAreaStyle Add()
    {
        XmlNode? formatNode = this.GetTopNode();
        ExcelPivotTableAreaStyle? s = new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, formatNode.FirstChild, this._pt);
        this._list.Add(s);

        return s;
    }

    /// <summary>
    /// Adds a pivot area style for the top right cells of the pivot table, to the right of any filter button, if reading order i set to Left-To-Right. 
    /// </summary>
    /// <returns></returns>
    public ExcelPivotTableAreaStyle AddTopEnd()
    {
        return this.AddTopEnd(null);
    }

    /// <summary>
    /// Adds a style for the top right cells of the pivot table, to the right of any filter button, if reading order i set to Left-To-Right. 
    /// </summary>
    /// <param name="offsetAddress">Offset address from the top-left cell to the right of any filter button. The top-left cell is refereced as A1. For example, B1:C1 will reference the second and third cell of the first row of the area. "null" will mean all cells</param>
    /// <returns></returns>
    public ExcelPivotTableAreaStyle AddTopEnd(string offsetAddress = null)
    {
        XmlNode? formatNode = this.GetTopNode();

        ExcelPivotTableAreaStyle? s = new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, formatNode.FirstChild, this._pt)
        {
            PivotAreaType = ePivotAreaType.TopEnd,
        };

        if (offsetAddress != null)
        {
            if (ExcelCellBase.IsSimpleAddress(offsetAddress) == false)
            {
                throw new ArgumentException("Offset address must be a valid address", "offsetAddress");
            }

            s.Offset = offsetAddress;
        }

        this._list.Add(s);

        return s;
    }

    /// <summary>
    /// Adds a style for the top left cells of the pivot table, if reading order i set to Left-To-Right
    /// </summary>
    /// <param name="offsetAddress">Offset address from the left cell. The top-left cell is refereced as A1. For example, B1:C1 will reference the second and third cell of the first row of the area. "null" will mean all cells </param>
    /// <returns></returns>
    public ExcelPivotTableAreaStyle AddTopStart(string offsetAddress = null)
    {
        XmlNode? formatNode = this.GetTopNode();

        ExcelPivotTableAreaStyle? s = new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, formatNode.FirstChild, this._pt)
        {
            PivotAreaType = ePivotAreaType.Origin, FieldIndex = 0, FieldPosition = 0, LabelOnly = true, DataOnly = false
        };

        if (offsetAddress != null)
        {
            if (ExcelCellBase.IsSimpleAddress(offsetAddress) == false)
            {
                throw new ArgumentException("Offset address must be a valid address", "offsetAddress");
            }

            s.Offset = offsetAddress;
        }

        this._list.Add(s);

        return s;
    }

    /// <summary>
    /// Adds a style for the filter box.
    /// </summary>
    /// <param name="field">The field with the box to style</param>
    /// <returns></returns>
    public ExcelPivotTableAreaStyle AddButtonField(ExcelPivotTableField field)
    {
        if (field == null)
        {
            throw new ArgumentException("Field can't be null");
        }

        XmlNode? formatNode = this.GetTopNode();

        ExcelPivotTableAreaStyle? s = new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, formatNode.FirstChild, this._pt)
        {
            PivotAreaType = ePivotAreaType.FieldButton, FieldIndex = field.Index, LabelOnly = true, DataOnly = false, Outline = false
        };

        if (field.IsColumnField)
        {
            s.Axis = ePivotTableAxis.ColumnAxis;
            s.FieldPosition = this._pt.ColumnFields.IndexOf(field);
        }
        else if (field.IsRowField)
        {
            s.Axis = ePivotTableAxis.RowAxis;
            s.FieldPosition = this._pt.RowFields.IndexOf(field);
        }
        else if (field.IsPageField)
        {
            s.Axis = ePivotTableAxis.PageAxis;
            s.FieldPosition = this._pt.PageFields.IndexOf(field);
        }

        this._list.Add(s);

        return s;
    }

    /// <summary>
    /// Adds a pivot area style that affects the whole table.
    /// </summary>
    /// <returns>The style object used to set the styles</returns>
    public ExcelPivotTableAreaStyle AddWholeTable()
    {
        return this.AddAll(false, false);
    }

    /// <summary>
    /// Adds a pivot area style that affects all labels
    /// </summary>
    /// <returns>The style object used to set the styles</returns>
    public ExcelPivotTableAreaStyle AddAllLabels()
    {
        return this.AddAll(true, false);
    }

    /// <summary>
    /// Adds a pivot area style that affects all data cells
    /// </summary>
    /// <returns>The style object used to set the styles</returns>
    public ExcelPivotTableAreaStyle AddAllData()
    {
        return this.AddAll(false, true);
    }

    internal ExcelPivotTableAreaStyle AddAll(bool labels, bool data)
    {
        XmlNode? formatNode = this.GetTopNode();

        ExcelPivotTableAreaStyle? s = new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, formatNode.FirstChild, this._pt)
        {
            PivotAreaType = ePivotAreaType.All, LabelOnly = labels, DataOnly = data
        };

        this._list.Add(s);

        return s;
    }

    /// <summary>
    /// Adds a style for the labels of a pivot table
    /// </summary>
    /// <param name="fields">The pivot table field that style affects</param>
    /// <returns></returns>
    public ExcelPivotTableAreaStyle AddLabel(params ExcelPivotTableField[] fields)
    {
        if (fields.Any(x => x == null))
        {
            throw new ArgumentException("Field in array can't be null");
        }

        ExcelPivotTableAreaStyle? s = this.Add();
        s.LabelOnly = true;
        s.FieldPosition = 0;
        s.Outline = false;

        foreach (ExcelPivotTableField? field in fields)
        {
            s.Conditions.Fields.Add(field);
        }

        return s;
    }

    /// <summary>
    /// Adds a style for the data area of a pivot table
    /// </summary>
    /// <param name="fields"></param>
    /// <returns></returns>
    public ExcelPivotTableAreaStyle AddData(params ExcelPivotTableField[] fields)
    {
        if (fields.Any(x => x == null))
        {
            throw new ArgumentException("Field in array can't be null");
        }

        ExcelPivotTableAreaStyle? s = this.Add();
        s.PivotAreaType = ePivotAreaType.Data;
        s.LabelOnly = false;
        s.FieldPosition = 0;
        s.Outline = false;

        foreach (ExcelPivotTableField? field in fields)
        {
            ExcelPivotAreaReference? r = s.Conditions.Fields.Add(this._pt, field.Index);
        }

        return s;
    }

    /// <summary>
    /// Adds a style for filter boxes.
    /// </summary>
    /// <param name="axis">The axis for the field buttons</param>
    /// <param name="index">The zero-based index in the axis collection</param>
    /// <returns></returns>
    public ExcelPivotTableAreaStyle AddButtonField(ePivotTableAxis axis, int index = 0)
    {
        if (index < 0)
        {
            throw new ArgumentException("Index must be positive", "index");
        }

        XmlNode? formatNode = this.GetTopNode();

        ExcelPivotTableAreaStyle? s = new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, formatNode.FirstChild, this._pt)
        {
            PivotAreaType = ePivotAreaType.FieldButton,
            FieldIndex = 0,
            FieldPosition = index,
            LabelOnly = true,
            DataOnly = false,
            Outline = false,
            Axis = axis
        };

        this._list.Add(s);

        return s;
    }

    internal ExcelPivotTableAreaStyle Add(ePivotAreaType type)
    {
        XmlNode? formatNode = this.GetTopNode();
        ExcelPivotTableAreaStyle? s = new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, formatNode.FirstChild, this._pt) { PivotAreaType = type };

        this._list.Add(s);

        return s;
    }

    internal ExcelPivotTableAreaStyle Add(ePivotAreaType type, ePivotTableAxis axis)
    {
        XmlNode? formatNode = this.GetTopNode();

        ExcelPivotTableAreaStyle? s =
            new ExcelPivotTableAreaStyle(this._styles.NameSpaceManager, formatNode.FirstChild, this._pt) { PivotAreaType = type, Axis = axis };

        this._list.Add(s);

        return s;
    }

    private XmlNode GetTopNode()
    {
        if (this._xmlHelper == null)
        {
            XmlNode? node = this._pt.CreateNode("d:formats");
            this._xmlHelper = XmlHelperFactory.Create(this._pt.NameSpaceManager, node);
        }

        XmlNode? retNode = this._xmlHelper.CreateNode("d:format", false, true);
        retNode.InnerXml = $"<pivotArea xmlns=\"{ExcelPackage.schemaMain}\"/>";

        return retNode;
    }
}