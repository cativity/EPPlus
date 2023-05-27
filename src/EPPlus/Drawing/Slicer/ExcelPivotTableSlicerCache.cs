/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/01/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/

using OfficeOpenXml.Constants;
using OfficeOpenXml.Filter;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer;

/// <summary>
/// Represents a pivot table slicer cache.
/// </summary>
public class ExcelPivotTableSlicerCache : ExcelSlicerCache
{
    internal ExcelPivotTableField _field = null;

    internal ExcelPivotTableSlicerCache(XmlNamespaceManager nameSpaceManager)
        : base(nameSpaceManager)
    {
        this.PivotTables = new ExcelSlicerPivotTableCollection(this);
    }

    internal void Init(ExcelWorkbook wb, string name, ExcelPivotTableField field)
    {
        if (wb._slicerCaches == null)
        {
            wb.LoadSlicerCaches();
        }

        this.CreatePart(wb);
        this.TopNode = this.SlicerCacheXml.DocumentElement;
        this.Name = "Slicer_" + ExcelAddressUtil.GetValidName(name);
        this._field = field;
        this.SourceName = this._field.Cache.Name;
        wb.Names.AddFormula(this.Name, "#N/A");
        this.PivotTables.Add(this._field._pivotTable);
        this.CreateWorkbookReference(wb, ExtLstUris.WorkbookSlicerPivotTableUri);
        this.SlicerCacheXml.Save(this.Part.GetStream());
        this.Data.Items.Refresh();
    }

    /// <summary>
    /// Init must be called before accessing any properties as it sets several properties.
    /// </summary>
    /// <param name="wb"></param>
    internal override void Init(ExcelWorkbook wb)
    {
        foreach (XmlElement ptElement in this.GetNodes("x14:pivotTables/x14:pivotTable"))
        {
            string? name = ptElement.GetAttribute("name");
            string? tabId = ptElement.GetAttribute("tabId");

            if (int.TryParse(tabId, out int sheetId))
            {
                ExcelWorksheet? ws = wb.Worksheets.GetBySheetID(sheetId);
                ExcelPivotTable? pt = ws?.PivotTables[name];

                if (pt != null)
                {
                    this._field ??= pt.Fields.Where(x => x.Cache.Name == this.SourceName).FirstOrDefault();

                    this.PivotTables._list.Add(pt);
                }
            }
        }
    }

    internal void Init(ExcelWorkbook wb, ExcelPivotTableSlicer slicer)
    {
        this._field = this.PivotTables[0].Fields.Where(x => x.Cache.Name == this.SourceName).FirstOrDefault();
        this.Init(wb);
    }

    /// <summary>
    /// The source type of the slicer
    /// </summary>
    public override eSlicerSourceType SourceType
    {
        get { return eSlicerSourceType.PivotTable; }
    }

    /// <summary>
    /// A collection of pivot tables attached to the slicer cache.
    /// </summary>
    public ExcelSlicerPivotTableCollection PivotTables { get; }

    ExcelPivotTableSlicerCacheTabularData _data = null;

    /// <summary>
    /// Tabular data for a pivot table slicer cache.
    /// </summary>
    public ExcelPivotTableSlicerCacheTabularData Data
    {
        get { return this._data ??= new ExcelPivotTableSlicerCacheTabularData(this.NameSpaceManager, this.TopNode, this); }
    }

    internal void UpdateItemsXml()
    {
        StringBuilder? sb = new StringBuilder();

        foreach (ExcelPivotTable? pt in this.PivotTables)
        {
            sb.Append($"<pivotTable name=\"{pt.Name}\" tabId=\"{pt.WorkSheet.SheetId}\"/>");
        }

        XmlNode? ptNode = this.CreateNode("x14:pivotTables");
        ptNode.InnerXml = sb.ToString();
        this.Data.UpdateItemsXml();
    }
}