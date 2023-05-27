/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/26/2020         EPPlus Software AB       EPPlus 5.4
 ******0*******************************************************************************************/

using OfficeOpenXml.Constants;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.IO;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer;

/// <summary>
/// Represents a pivot table slicer drawing object.
/// A pivot table slicer is attached to a pivot table fields item filter.
/// </summary>
public class ExcelPivotTableSlicer : ExcelSlicer<ExcelPivotTableSlicerCache>
{
    //internal ExcelPivotTableField _field;
    internal ExcelPivotTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelPivotTableField field, ExcelGroupShape parent = null)
        : base(drawings, node, parent)
    {
        this._ws = drawings.Worksheet;

        //_field = field;
        string? name = drawings.Worksheet.Workbook.GetSlicerName(field.Cache.Name);

        this.CreateDrawing(name);

        this.SlicerName = name;
        this.Caption = field.Name;
        this.RowHeight = 19;

        if (field.Slicer == null)
        {
            this.CacheName = "Slicer_" + ExcelAddressUtil.GetValidName(name);

            ExcelPivotTableSlicerCache? cache = new ExcelPivotTableSlicerCache(this.NameSpaceManager);
            field.Slicer ??= this;

            cache.Init(drawings.Worksheet.Workbook, name, field);
            this._cache = cache;
        }
        else
        {
            this.CacheName = field.Slicer.Cache.Name;
            this._cache = field.Slicer.Cache;
        }

        //If items has not been init, refresh!
        if (field._items == null)
        {
            field.Items.Refresh();
        }
    }

    internal ExcelPivotTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelPivotTableSlicerCache cache, ExcelGroupShape parent = null)
        : base(drawings, node, parent)
    {
        this._ws = drawings.Worksheet;
        string? name = drawings.Worksheet.Workbook.GetSlicerName(this.Cache.Name);
        this.CreateDrawing(name);

        this.SlicerName = name;
        this.Caption = name;
        this.RowHeight = 19;
        this.CacheName = this.Cache.Name;
        this._cache = cache;

        //If items has not been init, refresh!
        if (cache._field._items == null)
        {
            cache._field.Items.Refresh();
        }
    }

    internal ExcelPivotTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null)
        : base(drawings, node, parent)
    {
        this._ws = drawings.Worksheet;
        XmlNode? slicerNode = this._ws.SlicerXmlSources.GetSource(this.Name, eSlicerSourceType.PivotTable, out this._xmlSource);
        this._slicerXmlHelper = XmlHelperFactory.Create(this.NameSpaceManager, slicerNode);

        this._cache = drawings.Worksheet.Workbook.GetSlicerCaches(this.CacheName) as ExcelPivotTableSlicerCache;
    }

    private void CreateDrawing(string name)
    {
        XmlElement graphFrame = this.TopNode.OwnerDocument.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
        graphFrame.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
        graphFrame.SetAttribute("xmlns:a14", ExcelPackage.schemaDrawings2010);
        this.TopNode.AppendChild(graphFrame);

        graphFrame.InnerXml =
            string.Format(
                          "<mc:Choice Requires=\"a14\"><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"{2}\"><a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" name=\"{2}\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback xmlns=\"\"><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"{1}\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"12506325\" y=\"3238500\"/><a:ext cx=\"1828800\" cy=\"2524125\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"sv-SE\" sz=\"1100\"/><a:t>This shape represents a slicer. Slicers are supported in Excel 2010 or later. If the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2003 or earlier, the slicer cannot be used.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>",
                          this._id,
                          "{" + Guid.NewGuid().ToString() + "}",
                          name);

        this.TopNode.AppendChild(this.TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

        this._xmlSource = this._ws.SlicerXmlSources.GetOrCreateSource(eSlicerSourceType.Table);
        XmlElement? node = this._xmlSource.XmlDocument.CreateElement("slicer", ExcelPackage.schemaMainX14);
        this._xmlSource.XmlDocument.DocumentElement.AppendChild(node);
        this._slicerXmlHelper = XmlHelperFactory.Create(this.NameSpaceManager, node);

        XmlNode? extNode = this._ws.GetOrCreateExtLstSubNode(ExtLstUris.WorksheetSlicerPivotTableUri, "x14");

        if (extNode.InnerXml == "")
        {
            extNode.InnerXml = "<x14:slicerList/>";
            XmlNode? slNode = extNode.FirstChild;

            XmlHelper? xh = XmlHelperFactory.Create(this.NameSpaceManager, slNode);
            XmlElement? element = (XmlElement)xh.CreateNode("x14:slicer", false, true);
            element.SetAttribute("id", ExcelPackage.schemaRelationships, this._xmlSource.Rel.Id);
        }

        this.GetPositionSize();
    }

    internal override bool CheckSlicerNameIsUnique(string name)
    {
        return this._drawings.Worksheet.Workbook.CheckSlicerNameIsUnique(name);
    }

    internal override void DeleteMe()
    {
        this.Cache._field.Slicer = null;
        base.DeleteMe();
    }

    internal void CreateNewCache(ExcelPivotTableField field)
    {
        ExcelPivotTableSlicerCache? cache = new ExcelPivotTableSlicerCache(this._slicerXmlHelper.NameSpaceManager);
        cache.Init(this._ws.Workbook, this.SlicerName, field);
        this._cache = cache;
        this.CacheName = cache.Name;
    }
}