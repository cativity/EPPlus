﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/29/2020         EPPlus Software AB       EPPlus 5.3
 *************************************************************************************************/

using OfficeOpenXml.Constants;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer;

/// <summary>
/// Represents a slicer cache with a table as source
/// </summary>
public class ExcelTableSlicerCache : ExcelSlicerCache
{
    const string _extPath = "x14:extLst/d:ext";
    const string _topPath = _extPath + "/x15:tableSlicerCache";

    internal ExcelTableSlicerCache(XmlNamespaceManager nameSpaceManager)
        : base(nameSpaceManager)
    {
    }

    internal override void Init(ExcelWorkbook wb)
    {
        ExcelTable? tbl = wb.GetTable(this.TableId);
        this.TableColumn = tbl?.Columns.FirstOrDefault(x => x.Id == this.ColumnId);
        this.SlicerCacheXml = new XmlDocument();
        LoadXmlSafe(this.SlicerCacheXml, this.Part.GetStream());
    }

    internal void Init(ExcelTableColumn column, string cacheName)
    {
        this.TableColumn = column;
        ExcelWorkbook? wb = column.Table.WorkSheet.Workbook;
        this.CreatePart(wb);

        this.SlicerCacheXml.DocumentElement.InnerXml =
            $"<extLst><x:ext uri=\"{ExtLstUris.TableSlicerCacheUri}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:tableSlicerCache tableId=\"{column.Table.Id}\" column=\"{column.Id}\"/></x:ext></extLst>";

        this.TopNode = this.SlicerCacheXml.DocumentElement;
        this.Name = cacheName;
        this.SourceName = column.Name;
        this.SlicerCacheXml.Save(this.Part.GetStream(FileMode.Create, FileAccess.Write));
        this.CreateWorkbookReference(wb, ExtLstUris.WorkbookSlicerTableUri);
    }

    /// <summary>
    /// The source type for the slicer cache
    /// </summary>
    public override eSlicerSourceType SourceType => eSlicerSourceType.Table;

    /// <summary>
    /// The table column that is the source for the slicer
    /// </summary>
    public ExcelTableColumn TableColumn { get; private set; }

    const string _sortOrderPath = _topPath + "/@sortOrder";

    /// <summary>
    /// How the table slicer items are sorted
    /// </summary>
    public eSortOrder SortOrder
    {
        get => this.GetXmlNodeString(_sortOrderPath).ToEnum(eSortOrder.Ascending);
        set => this.SetXmlNodeString(_sortOrderPath, value.ToEnumString());
    }

    const string _crossFilterPath = _topPath + "/@crossFilter";

    /// <summary>
    /// How the items that are used in slicer cross filtering are displayed
    /// </summary>
    public eCrossFilter CrossFilter
    {
        get => this.GetXmlNodeString(_crossFilterPath).ToEnum(eCrossFilter.None);
        set => this.SetXmlNodeString(_crossFilterPath, value.ToEnumString());
    }

    const string _customListSortPath = _topPath + "/@customListSort";

    /// <summary>
    /// If custom lists are used when sorting the items
    /// </summary>
    public bool CustomListSort
    {
        get => this.GetXmlNodeBool(_customListSortPath, true);
        set => this.SetXmlNodeBool(_customListSortPath, value, true);
    }

    const string _hideItemsWithNoDataPath = "x15:slicerCacheHideItemsWithNoData";

    /// <summary>
    /// If true, items that have no data are not displayed
    /// </summary>
    public bool HideItemsWithNoData
    {
        get => this.ExistsNode(_extPath + "/" + _hideItemsWithNoDataPath);
        set
        {
            if (value)
            {
                XmlNode? node = this.CreateNode("x14:extLst/d:ext", false, true);
                ((XmlElement)node).SetAttribute("uri", "{470722E0-AACD-4C17-9CDC-17EF765DBC7E}");
                XmlHelper? helper = XmlHelperFactory.Create(this.NameSpaceManager, node);
                _ = helper.CreateNode(_hideItemsWithNoDataPath, false, true);
            }
            else
            {
                XmlNode? hideNode = this.GetNode(_extPath + "/" + _hideItemsWithNoDataPath);

                if (hideNode != null)
                {
                    _ = hideNode.ParentNode.ParentNode.RemoveChild(hideNode.ParentNode);
                }
            }
        }
    }

    const string _columnIndexPath = _topPath + "/@column";

    internal int ColumnId
    {
        get => this.GetXmlNodeInt(_columnIndexPath);
        set => this.SetXmlNodeInt(_columnIndexPath, value);
    }

    const string _tableIdPath = _topPath + "/@tableId";

    internal int TableId
    {
        get => this.GetXmlNodeInt(_tableIdPath);
        set => this.SetXmlNodeInt(_tableIdPath, value);
    }
}