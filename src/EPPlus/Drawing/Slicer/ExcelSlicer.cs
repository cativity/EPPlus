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

using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer;

/// <summary>
/// Base class for table and pivot table slicers.
/// </summary>
/// <typeparam name="T">The slicer cache data type</typeparam>
public abstract class ExcelSlicer<T> : ExcelDrawing
    where T : ExcelSlicerCache
{
    internal ExcelSlicerXmlSource _xmlSource;
    internal ExcelWorksheet _ws;
    internal XmlHelper _slicerXmlHelper;

    internal ExcelSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null)
        : base(drawings, node, "mc:AlternateContent/mc:Choice/xdr:graphicFrame", "xdr:nvGraphicFramePr/xdr:cNvPr", parent)
    {
        this._ws = drawings.Worksheet;
    }

    internal ExcelSlicer(ExcelDrawings drawings, XmlNode node, XmlDocument slicerXml, ExcelGroupShape parent = null)
        : base(drawings, node, "mc:AlternateContent/mc:Choice/xdr:graphicFrame", "xdr:nvGraphicFramePr/xdr:cNvPr", parent)
    {
        this._ws = drawings.Worksheet;
    }

    /// <summary>
    /// The type of drawing
    /// </summary>
    public override eDrawingType DrawingType
    {
        get { return eDrawingType.Slicer; }
    }

    /// <summary>
    /// The caption text of the slicer.
    /// </summary>
    public string Caption
    {
        get { return this._slicerXmlHelper.GetXmlNodeString("@caption"); }
        set { this._slicerXmlHelper.SetXmlNodeString("@caption", value); }
    }

    /// <summary>
    /// If the caption of the slicer is visible.
    /// </summary>
    public bool ShowCaption
    {
        get { return this._slicerXmlHelper.GetXmlNodeBool("@showCaption", true); }
        set { this._slicerXmlHelper.SetXmlNodeBool("@showCaption", value, true); }
    }

    /// <summary>
    /// The the name of the slicer.
    /// </summary>
    public string SlicerName
    {
        get { return this._slicerXmlHelper.GetXmlNodeString("@name"); }
        set
        {
            if (!this.CheckSlicerNameIsUnique(value))
            {
                if (this.Name != value)
                {
                    throw new InvalidOperationException("Slicer Name is not unique");
                }
            }

            if (this.Name != value)
            {
                this.Name = value;
            }

            this._slicerXmlHelper.SetXmlNodeString("@name", value);
        }
    }

    internal abstract bool CheckSlicerNameIsUnique(string name);

    /// <summary>
    /// Row height in points
    /// </summary>
    public double RowHeight
    {
        get { return this._slicerXmlHelper.GetXmlNodeEmuToPt("@rowHeight"); }
        set { this._slicerXmlHelper.SetXmlNodeEmuToPt("@rowHeight", value); }
    }

    /// <summary>
    /// The index of the starting item in the slicer. Default is 0.
    /// </summary>
    public int StartItem
    {
        get { return this._slicerXmlHelper.GetXmlNodeInt("@startItem", 0); }
        set { this._slicerXmlHelper.SetXmlNodeInt("@startItem", value, null, false); }
    }

    /// <summary>
    /// Number of columns. Default is 1.
    /// </summary>
    public int ColumnCount
    {
        get { return this._slicerXmlHelper.GetXmlNodeInt("@columnCount", 1); }
        set { this._slicerXmlHelper.SetXmlNodeInt("@columnCount", value, null, false); }
    }

    /// <summary>
    /// If the slicer view is locked or not.
    /// </summary>
    public bool LockedPosition
    {
        get { return this._slicerXmlHelper.GetXmlNodeBool("@lockedPosition", false); }
        set { this._slicerXmlHelper.SetXmlNodeBool("@lockedPosition", value, false); }
    }

    /// <summary>
    /// The build in slicer style.
    /// If set to Custom, the name in the <see cref="StyleName" /> is used 
    /// </summary>
    public eSlicerStyle Style
    {
        get { return this.StyleName.TranslateSlicerStyle(); }
        set
        {
            if (value == eSlicerStyle.None)
            {
                this.StyleName = "";
            }
            else if (value != eSlicerStyle.Custom)
            {
                this.StyleName = "SlicerStyle" + value.ToString();
            }
        }
    }

    /// <summary>
    /// The style name used for the slicer.
    /// <seealso cref="Style"/>
    /// </summary>
    public string StyleName
    {
        get { return this._slicerXmlHelper.GetXmlNodeString("@style"); }
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                this._slicerXmlHelper.DeleteNode("@style");

                return;
            }

            if (value.StartsWith("SlicerStyle", StringComparison.OrdinalIgnoreCase))
            {
                eSlicerStyle style = value.Substring(11).ToEnum(eSlicerStyle.Custom);

                if (style != eSlicerStyle.Custom || style != eSlicerStyle.None)
                {
                    this._slicerXmlHelper.SetXmlNodeString("@style", "SlicerStyle" + style);

                    return;
                }
            }

            this.Style = eSlicerStyle.Custom;
            this._slicerXmlHelper.SetXmlNodeString("@style", value);
        }
    }

    internal string CacheName
    {
        get { return this._slicerXmlHelper.GetXmlNodeString("@cache"); }
        set { this._slicerXmlHelper.SetXmlNodeString("@cache", value); }
    }

    internal ExcelSlicerCache _cache;

    /// <summary>
    /// A reference to the slicer cache.
    /// </summary>
    public T Cache
    {
        get
        {
            this._cache ??= this._drawings.Worksheet.Workbook.GetSlicerCaches(this.CacheName);

            return this._cache as T;
        }
    }

    internal override void DeleteMe()
    {
        if (this._slicerXmlHelper.TopNode.ParentNode.ChildNodes.Count == 1)
        {
            this._ws.RemoveSlicerReference(this._xmlSource);
            this._xmlSource = null;
        }

        _ = this._slicerXmlHelper.TopNode.ParentNode.RemoveChild(this._slicerXmlHelper.TopNode);

        this._ws.Workbook.RemoveSlicerCacheReference(this.Cache.CacheRel.Id, this.Cache.SourceType);
        this._ws.Workbook.Names.Remove(this.Name);

        if (this.Cache.Part.Package.PartExists(this.Cache.Uri))
        {
            this._drawings.Worksheet.Workbook._package.ZipPackage.DeletePart(this.Cache.Uri);
        }

        base.DeleteMe();
    }
}