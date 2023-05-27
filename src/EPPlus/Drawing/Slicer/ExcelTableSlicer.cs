/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/26/2020         EPPlus Software AB       EPPlus 5.3
 ******0*******************************************************************************************/

using OfficeOpenXml.Constants;
using OfficeOpenXml.Filter;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer;

/// <summary>
/// Represents a table slicer drawing object.
/// A table slicer is attached to a table column value filter.
/// </summary>
public class ExcelTableSlicer : ExcelSlicer<ExcelTableSlicerCache>
{
    internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null)
        : base(drawings, node, parent)
    {
        this._ws = drawings.Worksheet;
        XmlNode? slicerNode = this._ws.SlicerXmlSources.GetSource(this.Name, eSlicerSourceType.Table, out this._xmlSource);
        this._slicerXmlHelper = XmlHelperFactory.Create(this.NameSpaceManager, slicerNode);

        _ = this._ws.Workbook.SlicerCaches.TryGetValue(this.CacheName, out ExcelSlicerCache cache);
        this._cache = (ExcelTableSlicerCache)cache;

        this.TableColumn = this.GetTableColumn();
    }

    internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelTableColumn column)
        : base(drawings, node)
    {
        this.TableColumn = column;
        column.Slicer ??= this;

        string? name = drawings.Worksheet.Workbook.GetSlicerName(column.Name);
        this.CreateDrawing(name);
        this.SlicerName = name;

        this.Caption = column.Name;
        this.RowHeight = 19;
        this.CacheName = "Slicer_" + ExcelAddressUtil.GetValidName(name);

        ExcelTableSlicerCache? cache = new ExcelTableSlicerCache(this.NameSpaceManager);
        cache.Init(column, this.CacheName);
        this._cache = cache;
    }

    private ExcelTableColumn GetTableColumn()
    {
        foreach (ExcelWorksheet? ws in this._drawings.Worksheet.Workbook.Worksheets)
        {
            foreach (ExcelTable? t in ws.Tables)
            {
                if (t.Id == this.Cache.TableId)
                {
                    return t.Columns.Where(x => x.Id == this.Cache.ColumnId).Single();
                }
            }
        }

        return null;
    }

    internal override void DeleteMe()
    {
        try
        {
            this.TableColumn.Slicer = null;
        }
        catch (Exception ex)
        {
            throw new InvalidDataException("EPPlus internal error when deleting the slicer.", ex);
        }

        base.DeleteMe();
    }

    private void CreateDrawing(string name)
    {
        XmlElement graphFrame = this.TopNode.OwnerDocument.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
        graphFrame.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
        graphFrame.SetAttribute("xmlns:sle15", ExcelPackage.schemaSlicer);
        _ = this.TopNode.AppendChild(graphFrame);

        graphFrame.InnerXml =
            string.Format(
                          "<mc:Choice Requires=\"sle15\"><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"{2}\"><a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" name=\"{2}\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback xmlns=\"\"><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"{2}\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"1200150\" y=\"2971800\"/><a:ext cx=\"1828800\" cy=\"2524125\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\"/><a:t>This shape represents a table slicer. Table slicers are not supported in this version of Excel.If the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2007 or earlier, the slicer can't be used.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>",
                          this._id,
                          "{" + Guid.NewGuid().ToString() + "}",
                          name);

        _ = this.TopNode.AppendChild(this.TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

        this._xmlSource = this._ws.SlicerXmlSources.GetOrCreateSource(eSlicerSourceType.Table);
        XmlElement? node = this._xmlSource.XmlDocument.CreateElement("slicer", ExcelPackage.schemaMainX14);
        _ = this._xmlSource.XmlDocument.DocumentElement.AppendChild(node);
        this._slicerXmlHelper = XmlHelperFactory.Create(this.NameSpaceManager, node);

        XmlNode? extNode = this._ws.GetOrCreateExtLstSubNode(ExtLstUris.WorksheetSlicerTableUri, "x14");

        if (extNode.InnerXml == "")
        {
            extNode.InnerXml = "<x14:slicerList/>";
            XmlHelper? xh = XmlHelperFactory.Create(this.NameSpaceManager, extNode.FirstChild);
            XmlElement? element = (XmlElement)xh.CreateNode("x14:slicer", false, true);
            _ = element.SetAttribute("id", ExcelPackage.schemaRelationships, this._xmlSource.Rel.Id);
        }

        this.GetPositionSize();
    }

    /// <summary>
    /// The table column that the slicer is connected to.
    /// </summary>
    public ExcelTableColumn TableColumn { get; internal set; }

    /// <summary>
    /// The value filters for the slicer. This is the same filter as the filter for the table.
    /// This filter is a value filter.
    /// </summary>
    public ExcelValueFilterCollection FilterValues
    {
        get
        {
            ExcelValueFilterColumn? f = this.TableColumn.Table.AutoFilter.Columns[this.TableColumn.Position] as ExcelValueFilterColumn;

            if (f != null)
            {
                return f.Filters;
            }
            else
            {
                return null;
            }
        }
    }

    internal override bool CheckSlicerNameIsUnique(string name)
    {
        return this._drawings.Worksheet.Workbook.CheckSlicerNameIsUnique(name);
    }

    internal void CreateNewCache()
    {
        ExcelTableSlicerCache? cache = new ExcelTableSlicerCache(this._slicerXmlHelper.NameSpaceManager);
        cache.Init(this.Cache.TableColumn, "Slicer_" + this.SlicerName);
        this._cache = cache;
        this.CacheName = cache.Name;
    }
}