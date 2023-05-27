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
using System.Linq;
using System.Xml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.Utils.TypeConversion;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Base class for drawings. 
/// Drawings are Charts, Shapes and Pictures.
/// </summary>
public class ExcelDrawing : XmlHelper, IDisposable
{
    internal ExcelDrawings _drawings;
    internal ExcelGroupShape _parent;
    internal string _topPath, _nvPrPath, _hyperLinkPath;
    internal string _topPathUngrouped, _nvPrPathUngrouped;
    internal int _id;
    internal const float STANDARD_DPI = 96;
    /// <summary>
    /// The ratio between EMU and Pixels
    /// </summary>
    public const int EMU_PER_PIXEL = 9525;
    /// <summary>
    /// The ratio between EMU and Points
    /// </summary>
    public const int EMU_PER_POINT = 12700;
    /// <summary>
    /// The ratio between EMU and centimeters
    /// </summary>
    public const int EMU_PER_CM = 360000;
    /// <summary>
    /// The ratio between EMU and milimeters
    /// </summary>
    public const int EMU_PER_MM = 3600000;
    /// <summary>
    /// The ratio between EMU and US Inches
    /// </summary>
    public const int EMU_PER_US_INCH = 914400;
    /// <summary>
    /// The ratio between EMU and pica
    /// </summary>
    public const int EMU_PER_PICA = EMU_PER_US_INCH / 6;

    internal double _width = double.MinValue, _height = double.MinValue, _top = double.MinValue, _left = double.MinValue;
    internal static readonly string[] _schemaNodeOrderSpPr = new string[] { "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "pattFill", "grpFill", "blipFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d" };

    internal bool _doNotAdjust = false;
    internal ExcelDrawing(ExcelDrawings drawings, XmlNode node, string topPath, string nvPrPath, ExcelGroupShape parent = null) :
        base(drawings.NameSpaceManager, node)
    {
        this._drawings = drawings;
        this._parent = parent;
        if (node != null)   //No drawing, chart xml only. This currently happends when created from a chart template
        {
            this.TopNode = node;
                
            if(this.DrawingType==eDrawingType.Control || drawings.Worksheet.Workbook._nextDrawingId >= 1025)
            {
                this._id = drawings.Worksheet._nextControlId++;
            }
            else
            {
                this._id = drawings.Worksheet.Workbook._nextDrawingId++;
            }

            this.AddSchemaNodeOrder(new string[] { "from", "pos", "to", "ext", "pic", "graphicFrame", "sp", "cxnSp ","grpSp", "nvSpPr", "nvCxnSpPr", "nvGraphicFramePr", "spPr", "style", "AlternateContent", "clientData" }, _schemaNodeOrderSpPr);
            this._topPathUngrouped = topPath;
            this._nvPrPathUngrouped = nvPrPath;
            if (this._parent == null)
            {
                this.AdjustXPathsForGrouping(false);
                this.CellAnchor = GetAnchorFromName(node.LocalName);
                this.SetPositionProperties(drawings, node);
                this.GetPositionSize();                                  //Get the drawing position and size, so we can adjust it upon save, if the normal font is changed 

                string relID = this.GetXmlNodeString(this._hyperLinkPath + "/@r:id");
                if (!string.IsNullOrEmpty(relID))
                {
                    this.HypRel = drawings.Part.GetRelationship(relID);
                    if (this.HypRel.TargetUri == null)
                    {
                        if (!string.IsNullOrEmpty(this.HypRel.Target))
                        {
                            this._hyperLink = new ExcelHyperLink(this.HypRel.Target.Substring(1), "");
                        }
                    }
                    else
                    {
                        if (this.HypRel.TargetUri.IsAbsoluteUri)
                        {
                            this._hyperLink = new ExcelHyperLink(this.HypRel.TargetUri.AbsoluteUri);
                        }
                        else
                        {
                            this._hyperLink = new ExcelHyperLink(this.HypRel.TargetUri.OriginalString, UriKind.Relative);
                        }
                    }
                    if (this.Hyperlink is ExcelHyperLink ehl)
                    {
                        ehl.ToolTip = this.GetXmlNodeString(this._hyperLinkPath + "/@tooltip");
                    }
                }
            }
            else
            {
                this.AdjustXPathsForGrouping(true);
                this.SetPositionProperties(drawings, node);
                this.GetPositionSize();                                  //Get the drawing position and size, so we can adjust it upon save, if the normal font is changed 
            }
        }   
    }

    internal virtual void AdjustXPathsForGrouping(bool group)
    {
        if(group)
        {
            this._topPath = this._topPathUngrouped.IndexOf('/') > 0 ? this._topPathUngrouped.Substring(this._topPathUngrouped.IndexOf('/')+1) : "";
            if(this._topPath=="")
            {
                this._nvPrPath = this._nvPrPathUngrouped;
            }
            else
            {
                this._nvPrPath = this._topPath + "/" + this._nvPrPathUngrouped;
            }
        }
        else
        {
            this._topPath = this._topPathUngrouped;
            this._nvPrPath = this._topPath + "/" + this._nvPrPathUngrouped;
        }

        this._hyperLinkPath = $"{this._nvPrPath}/a:hlinkClick";
    }

    internal void SetGroupChild(XmlNode offNode, XmlNode extNode)
    {
        this.CellAnchor = eEditAs.Absolute;

        this.From = null;
        this.To = null;
        this.Position = new ExcelDrawingCoordinate(this.NameSpaceManager, offNode, this.GetPositionSize);
        this.Size = new ExcelDrawingSize(this.NameSpaceManager, extNode, this.GetPositionSize);
    }

    private void SetPositionProperties(ExcelDrawings drawings, XmlNode node)
    {
        if (this._parent == null) //Top level drawing
        {
            XmlNode posNode = node.SelectSingleNode("xdr:from", drawings.NameSpaceManager);
            if (posNode != null)
            {
                this.From = new ExcelPosition(drawings.NameSpaceManager, posNode, this.GetPositionSize);
            }
            else
            {
                posNode = node.SelectSingleNode("xdr:pos", drawings.NameSpaceManager);
                if (posNode != null)
                {
                    this.Position = new ExcelDrawingCoordinate(drawings.NameSpaceManager, posNode, this.GetPositionSize);
                }
            }
            posNode = node.SelectSingleNode("xdr:to", drawings.NameSpaceManager);
            if (posNode != null)
            {
                this.To = new ExcelPosition(drawings.NameSpaceManager, posNode, this.GetPositionSize);
            }
            else
            {
                this.To = null;
                posNode = node.SelectSingleNode("xdr:ext", drawings.NameSpaceManager);
                if (posNode != null)
                {
                    this.Size = new ExcelDrawingSize(drawings.NameSpaceManager, posNode, this.GetPositionSize);
                }
            }
        }
        else //Child to Group shape
        {
            this.From = null;
            this.To = null;
            XmlNode posNode = this.GetXFrameNode(node, "a:off");
            if (posNode != null)
            {
                this.Position = new ExcelDrawingCoordinate(drawings.NameSpaceManager, posNode, this.GetPositionSize);
            }

            posNode = this.GetXFrameNode(node, "a:ext");
            if (posNode != null)
            {
                this.Size = new ExcelDrawingSize(drawings.NameSpaceManager, posNode, this.GetPositionSize);
            }
        }
    }

    private XmlNode GetXFrameNode(XmlNode node, string child)
    {
        if(node.LocalName == "AlternateContent")
        {
            node = node.GetChildAtPosition(0).GetChildAtPosition(0);
        }
        if (node.LocalName == "grpSp")
        {
            return node.SelectSingleNode($"xdr:grpSpPr/a:xfrm/{child}", this.NameSpaceManager);
        }
        else if (node.LocalName == "graphicFrame")
        {
            return node.SelectSingleNode($"xdr:xfrm/{child}", this.NameSpaceManager);
        }
        else
        {
            return node.SelectSingleNode($"xdr:spPr/a:xfrm/{child}", this.NameSpaceManager);
        }
    }

    internal bool IsWithinColumnRange(int colFrom, int colTo)
    {
        if (this.CellAnchor == eEditAs.OneCell)
        {
            this.GetToColumnFromPixels(this._width, out int col, out _);
            return (this.From.Column > colFrom - 1 || (this.From.Column == colFrom - 1 && this.From.ColumnOff == 0)) && col <= colTo;
        }
        else if (this.CellAnchor == eEditAs.TwoCell)
        {
            return (this.From.Column > colFrom - 1 || (this.From.Column == colFrom - 1 && this.From.ColumnOff == 0)) && this.To.Column <= colTo;
        }
        else
        {
            return false;
        }
    }
    internal bool IsWithinRowRange(int rowFrom, int rowTo)
    {
        if (this.CellAnchor == eEditAs.OneCell)
        {
            this.GetToRowFromPixels(this._height, out int row, out int pixOff);
            return (this.From.Row > rowFrom - 1 || (this.From.Row == rowFrom - 1 && this.From.RowOff == 0)) && row <= rowTo;
        }
        else if (this.CellAnchor == eEditAs.TwoCell)
        {
            return (this.From.Row > rowFrom - 1 || (this.From.Row == rowFrom - 1 && this.From.RowOff == 0)) && this.To.Row <= rowTo;
        }
        else
        {
            return false;
        }
    }

    internal static eEditAs GetAnchorFromName(string topElementName)
    {
        switch (topElementName)
        {
            case "oneCellAnchor":
                return eEditAs.OneCell;
            case "absoluteAnchor":
                return eEditAs.Absolute;
            default:
                return eEditAs.TwoCell;
        }
    }
    /// <summary>
    /// The type of drawing
    /// </summary>
    public virtual eDrawingType DrawingType
    {
        get
        {
            return eDrawingType.Drawing;
        }
    }
    /// <summary>
    /// The name of the drawing object
    /// </summary>
    public virtual string Name 
    {
        get
        {
            try
            {
                if (this._nvPrPath == "")
                {
                    return "";
                }

                return this.GetXmlNodeString(this._nvPrPath+"/@name");
            }
            catch
            {
                return ""; 
            }
        }
        set
        {
            try
            {
                if (this._nvPrPath == "")
                {
                    throw new NotImplementedException();
                }

                this.SetXmlNodeString(this._nvPrPath + "/@name", value);
                if (this is ExcelSlicer<ExcelTableSlicerCache> ts)
                {
                    this.SetXmlNodeString(this._nvPrPath + "/../../a:graphic/a:graphicData/sle:slicer/@name", value);
                    ts.SlicerName = value;
                }
                else if (this is ExcelSlicer<ExcelPivotTableSlicerCache> pts)
                {
                    this.SetXmlNodeString(this._nvPrPath + "/../../a:graphic/a:graphicData/sle:slicer/@name", value);
                    pts.SlicerName = value;
                }
            }
            catch
            {
                throw new NotImplementedException();
            }
        }
    }


    /// <summary>
    /// A description of the drawing object
    /// </summary>
    public string Description
    {
        get
        {
            try
            {
                if (this._nvPrPath == "")
                {
                    return "";
                }

                return this.GetXmlNodeString(this._nvPrPath + "/@descr");
            }
            catch
            {
                return "";
            }
        }
        set
        {
            try
            {
                if (this._nvPrPath == "")
                {
                    throw new NotImplementedException();
                }

                this.SetXmlNodeString(this._nvPrPath + "/@descr", value);
            }
            catch
            {
                throw new NotImplementedException();
            }
        }
    }
    /// <summary>
    /// How Excel resize drawings when the column width is changed within Excel.
    /// </summary>
    public eEditAs EditAs
    {
        get
        {
            try
            {
                if (this._parent!=null && this.DrawingType == eDrawingType.Control)
                {
                    return ((ExcelControl)this).GetCellAnchorFromWorksheetXml();
                }
                if (this.CellAnchor == eEditAs.TwoCell)
                {
                    string s = this.GetXmlNodeString("@editAs");
                    if (s == "")
                    {
                        return eEditAs.TwoCell;
                    }
                    else
                    {
                        return (eEditAs)Enum.Parse(typeof(eEditAs), s, true);
                    }
                }
                else
                {
                    return this.CellAnchor;
                }
            }
            catch
            {
                return eEditAs.TwoCell;
            }
        }
        set
        {
            if(this._parent!=null)
            {
                if(this.DrawingType==eDrawingType.Control)
                {
                    ((ExcelControl)this).SetCellAnchor(value);
                }
                else
                {
                    throw new InvalidOperationException("EditAs can't be set when a drawing is a part of a group.");
                }
            }
            else if (this.CellAnchor == eEditAs.TwoCell)
            {
                string s = value.ToString();
                this.SetXmlNodeString("@editAs", s.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1, s.Length - 1));
            }
            else if(this.CellAnchor!=value)
            {
                throw new InvalidOperationException("EditAs can only be set when CellAnchor is set to TwoCellAnchor");
            }
        }
    }
    const string lockedPath="xdr:clientData/@fLocksWithSheet";
    /// <summary>
    /// Lock drawing
    /// </summary>
    public virtual bool Locked
    {
        get
        {
            return this.GetXmlNodeBool(lockedPath, true);
        }
        set
        {
            this.SetXmlNodeBool(lockedPath, value);
        }
    }
    const string printPath = "xdr:clientData/@fPrintsWithSheet";
    /// <summary>
    /// Print drawing with sheet
    /// </summary>
    public virtual bool Print
    {
        get
        {
            return this.GetXmlNodeBool(printPath, true);
        }
        set
        {
            this.SetXmlNodeBool(printPath, value);
        }
    }
    /// <summary>
    /// Top Left position, if the shape is of the one- or two- cell anchor type
    /// Otherwise this propery is set to null
    /// </summary>
    public ExcelPosition From
    {
        get;
        private set;
    }
    /// <summary>
    /// Top Left position, if the shape is of the absolute anchor type
    /// </summary>
    public ExcelDrawingCoordinate Position
    {
        get;
        private set;
    }
    /// <summary>
    /// The extent of the shape, if the shape is of the one- or absolute- anchor type.
    /// Otherwise this propery is set to null
    /// </summary>
    public ExcelDrawingSize Size
    {
        get;
        private set;
    }
    /// <summary>
    /// Bottom right position
    /// </summary>
    public ExcelPosition To { get; private set; } = null;
    Uri _hyperLink=null;
    /// <summary>
    /// Hyperlink
    /// </summary>
    public Uri Hyperlink
    {
        get
        {
            return this._hyperLink;
        }
        set
        {
            if (this._hyperLink != null)
            {
                this.DeleteNode(this._hyperLinkPath);
                if (this.HypRel != null)
                {
                    this._drawings._package.ZipPackage.DeletePart(UriHelper.ResolvePartUri(this.HypRel.SourceUri, this.HypRel.TargetUri));
                }
            }

            if (value != null)
            {
                if(value is ExcelHyperLink el && !string.IsNullOrEmpty(el.ReferenceAddress))
                {
                    this.HypRel = this._drawings.Part.CreateRelationship("#" + new ExcelAddress(el.ReferenceAddress).FullAddress, TargetMode.Internal, ExcelPackage.schemaHyperlink);
                }
                else
                {
                    this.HypRel = this._drawings.Part.CreateRelationship(value, TargetMode.External, ExcelPackage.schemaHyperlink);
                }

                this.SetXmlNodeString(this._hyperLinkPath + "/@r:id", this.HypRel.Id);
                if (this.Hyperlink is ExcelHyperLink excelLink)
                {
                    this.SetXmlNodeString(this._hyperLinkPath + "/@tooltip", excelLink.ToolTip);
                }
            }

            this._hyperLink = value;
        }
    }
    ExcelDrawingAsType _as = null;
    /// <summary>
    /// Provides access to type conversion for all top-level drawing classes.
    /// </summary>
    public ExcelDrawingAsType As
    {
        get { return this._as ??= new ExcelDrawingAsType(this); }
    }
    internal ZipPackageRelationship HypRel { get; set; }
    /// <summary>
    /// Add new Drawing types here
    /// </summary>
    /// <param name="drawings">The drawing collection</param>
    /// <param name="node">Xml top node</param>
    /// <returns>The Drawing object</returns>
    internal static ExcelDrawing GetDrawing(ExcelDrawings drawings, XmlNode node)
    {
        if (node.ChildNodes.Count < 3)
        {
            return null; //Invalid formatted anchor node, ignore
        }

        XmlElement drawNode = (XmlElement)node.GetChildAtPosition(2);
        return GetDrawingFromNode(drawings, node, drawNode);
    }

    internal static ExcelDrawing GetDrawingFromNode(ExcelDrawings drawings, XmlNode node, XmlElement drawNode, ExcelGroupShape parent=null)
    {
        switch (drawNode.LocalName)
        {
            case "sp":
                return GetShapeOrControl(drawings, node, drawNode, parent);
            case "pic":
                return new ExcelPicture(drawings, node, parent);
            case "graphicFrame":
                return ExcelChart.GetChart(drawings, node, parent);
            case "grpSp":
                return new ExcelGroupShape(drawings, node, parent);
            case "cxnSp":
                return new ExcelConnectionShape(drawings, node, parent);
            case "contentPart":
                //Not handled yet, return as standard drawing below
                break;
            case "AlternateContent":
                XmlElement choice = drawNode.FirstChild as XmlElement;
                if (choice != null && choice.LocalName == "Choice")
                {
                    string? req = choice.GetAttribute("Requires");  //NOTE:Can be space sparated. Might have to implement functinality for this.
                    string? ns = drawNode.GetAttribute($"xmlns:{req}");
                    if (ns == "")
                    {
                        ns = choice.GetAttribute($"xmlns:{req}");
                    }
                    switch (ns)
                    {
                        case ExcelPackage.schemaChartEx2015_9_8:
                        case ExcelPackage.schemaChartEx2015_10_21:
                        case ExcelPackage.schemaChartEx2016_5_10:
                            return ExcelChart.GetChartEx(drawings, node, parent);
                        case ExcelPackage.schemaSlicer:
                            return new ExcelTableSlicer(drawings, node, parent);
                        case ExcelPackage.schemaDrawings2010:
                            if (choice.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/@uri", drawings.NameSpaceManager)?.Value == ExcelPackage.schemaSlicer2010)
                            {
                                return new ExcelPivotTableSlicer(drawings, node, parent);
                            }
                            else if (choice.ChildNodes.Count > 0 && choice.FirstChild.LocalName=="sp")
                            {
                                return GetShapeOrControl(drawings, node, (XmlElement)choice.FirstChild, parent);
                            }
                            break;

                    }
                }
                break;
        }
        return new ExcelDrawing(drawings, node, "", "");
    }

    private static ExcelDrawing GetShapeOrControl(ExcelDrawings drawings, XmlNode node, XmlElement drawNode, ExcelGroupShape parent)
    {
        int shapeId = GetControlShapeId(drawNode, drawings.NameSpaceManager);
        ControlInternal? control = drawings.Worksheet.Controls.GetControlByShapeId(shapeId);
        if (control != null)
        {
            return ControlFactory.GetControl(drawings, drawNode, control, parent);
        }
        else
        {
            return new ExcelShape(drawings, node, parent);
        }
    }
            
    private static int GetControlShapeId(XmlElement drawNode, XmlNamespaceManager nameSpaceManager)
    {
        XmlNode? idNode = drawNode.SelectSingleNode("xdr:nvSpPr/xdr:cNvPr/@id", nameSpaceManager);
        if(idNode!=null)
        {
            return int.Parse(idNode.Value);
        }
        return -1;
    }

    internal int Id
    {
        get { return this._id; }
    }
    #region "Internal sizing functions"
    internal void GetFromBounds(out int fromRow, out int fromRowOff, out int fromCol, out int fromColOff)
    {
        if (this.CellAnchor == eEditAs.Absolute)
        {
            this.GetToRowFromPixels(this.Position.Y, out fromRow, out fromRowOff);
            this.GetToColumnFromPixels(this.Position.X, out fromCol, out fromColOff);
        }
        else
        {
            fromRow = this.From.Row;
            fromRowOff = this.From.RowOff;
            fromCol = this.From.Column;
            fromColOff = this.From.ColumnOff;
        }
    }
    internal void GetToBounds(out int toRow, out int toRowOff, out int toCol, out int toColOff)
    {
        if (this.CellAnchor == eEditAs.Absolute)
        {
            this.GetToRowFromPixels((this.Position.Y + this.Size.Height) / EMU_PER_PIXEL, out toRow, out toRowOff);
            this.GetToColumnFromPixels(this.Position.X + (this.Size.Width / EMU_PER_PIXEL), out toCol, out toColOff);
        }
        else
        {
            if (this.CellAnchor == eEditAs.TwoCell)
            {
                toRow = this.To.Row;
                toRowOff = this.To.RowOff;
                toCol = this.To.Column;
                toColOff = this.To.ColumnOff;
            }
            else
            {
                this.GetToRowFromPixels(this.Size.Height / EMU_PER_PIXEL, out toRow, out toRowOff, this.From.Row, this.From.RowOff);
                this.GetToColumnFromPixels(this.Size.Width / EMU_PER_PIXEL, out toCol, out toColOff, this.From.Column, this.From.ColumnOff);
            }
        }
    }
    internal int GetPixelLeft()
    {
        int pix;
        if (this.CellAnchor == eEditAs.Absolute)
        {
            pix = this.Position.X / EMU_PER_PIXEL;
        }
        else
        {
            ExcelWorksheet ws = this._drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;

            pix = 0;
            for (int col = 0; col < this.From.Column; col++)
            {
                pix += ws.GetColumnWidthPixels(col, mdw);
            }
            pix += this.From.ColumnOff / EMU_PER_PIXEL;
        }

        return pix;
    }
    internal int GetPixelTop()
    {
        int pix;
        if (this.CellAnchor == eEditAs.Absolute)
        {
            pix = this.Position.Y / EMU_PER_PIXEL;
        }
        else
        {
            pix = 0;
            Dictionary<int, double>? cache = this._drawings.Worksheet.RowHeightCache;
            for (int row = 0; row < this.From.Row; row++)
            {
                lock (cache)
                {
                    if (!cache.ContainsKey(row))
                    {
                        cache.Add(row, this._drawings.Worksheet.GetRowHeight(row + 1));
                    }
                }
                pix += (int)(cache[row] / 0.75);
            }
            pix += this.From.RowOff / EMU_PER_PIXEL;
        }
        return pix;
    }
    internal double GetPixelWidth()
    {
        double pix;
        if (this.CellAnchor == eEditAs.TwoCell)
        {
            ExcelWorksheet ws = this._drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;

            pix = -this.From.ColumnOff / (double)EMU_PER_PIXEL;
            for (int col = this.From.Column + 1; col <= this.To.Column; col++)
            {
                pix += (double)decimal.Truncate(((256 * ws.GetColumnWidth(col)) + decimal.Truncate(128 / (decimal)mdw)) / 256 * mdw);
            }

            double w = (double)decimal.Truncate(((256 * ws.GetColumnWidth(this.To.Column + 1)) + decimal.Truncate(128 / (decimal)mdw)) / 256 * mdw);
            pix += Math.Min(w, Convert.ToDouble(this.To.ColumnOff) / EMU_PER_PIXEL);
        }
        else
        {
            pix = this.Size.Width / (double)EMU_PER_PIXEL;
        }
        return pix;
    }
    internal double GetPixelHeight()
    {
        double pix;
        if (this.CellAnchor == eEditAs.TwoCell)
        {
            ExcelWorksheet ws = this._drawings.Worksheet;

            pix = -(this.From.RowOff / (double)EMU_PER_PIXEL);
            for (int row = this.From.Row + 1; row <= this.To.Row; row++)
            {
                pix += ws.GetRowHeight(row) / 0.75;
            }
            double h = ws.GetRowHeight(this.To.Row + 1) / 0.75;
            pix += Math.Min(h, Convert.ToDouble(this.To.RowOff) / EMU_PER_PIXEL);
        }
        else
        {
            pix = this.Size.Height / (double)EMU_PER_PIXEL;
        }
        return pix;
    }

    internal void SetPixelTop(double pixels)
    {
        this._doNotAdjust = true;
        if (this.CellAnchor == eEditAs.Absolute)
        {
            this.Position.Y = (int)(pixels * EMU_PER_PIXEL);
        }
        else
        {
            this.CalcRowFromPixelTop(pixels, out int row, out int rowOff);
            this.From.Row = row;
            this.From.RowOff = rowOff;
        }

        this._top = pixels;
        this._doNotAdjust = false;
    }

    internal void CalcRowFromPixelTop(double pixels, out int row, out int rowOff)
    {
        ExcelWorksheet ws = this._drawings.Worksheet;
        decimal mdw = ws.Workbook.MaxFontWidth;
        double prevPix = 0;
        double pix = ws.GetRowHeight(1) / 0.75;
        int r = 2;
        while (pix < pixels)
        {
            prevPix = pix;
            pix += (int)(ws.GetRowHeight(r++) / 0.75);
        }

        if (pix == pixels)
        {
            row = r - 1;
            rowOff = 0;
        }
        else
        {
            row = r - 2;
            rowOff = (int)(pixels - prevPix) * EMU_PER_PIXEL;
        }
    }

    internal void SetPixelLeft(double pixels)
    {
        this._doNotAdjust = true;
        if (this.CellAnchor == eEditAs.Absolute)
        {
            this.Position.X = (int)(pixels * EMU_PER_PIXEL);
        }
        else
        {
            this.CalcColFromPixelLeft(pixels, out int col, out int colOff);
            this.From.Column = col;
            this.From.ColumnOff = colOff;
        }

        this._doNotAdjust = false;

        this._left = pixels;
    }
    internal void CalcColFromPixelLeft(double pixels, out int column, out int columnOff)
    {

        ExcelWorksheet ws = this._drawings.Worksheet;
        decimal mdw = ws.Workbook.MaxFontWidth;
        double prevPix = 0;
        double pix = (int)decimal.Truncate(((256 * ws.GetColumnWidth(1)) + decimal.Truncate(128 / (decimal)mdw)) / 256 * mdw);
        int col = 2;

        while (pix < pixels)
        {
            prevPix = pix;
            pix += (int)decimal.Truncate(((256 * ws.GetColumnWidth(col++)) + decimal.Truncate(128 / (decimal)mdw)) / 256 * mdw);
        }
        if (pix == pixels)
        {
            column = col - 1;
            columnOff = 0;
        }
        else
        {
            column = col - 2;
            columnOff = (int)(pixels - prevPix) * EMU_PER_PIXEL;
        }
    }
    internal void SetPixelHeight(double pixels)
    {
        if (this.CellAnchor == eEditAs.TwoCell)
        {
            this._doNotAdjust = true;
            this.GetToRowFromPixels(pixels,  out int toRow, out int pixOff);
            this.To.Row = toRow;
            this.To.RowOff = pixOff;
            this._doNotAdjust = false;
        }
        else
        {
            this.Size.Height = (long)Math.Round(pixels * EMU_PER_PIXEL);
        }
    }

    internal void GetToRowFromPixels(double pixels, out int toRow, out int rowOff, int fromRow=-1, int fromRowOff=-1)
    {
        if(fromRow<0)
        {
            fromRow = this.From.Row;
            fromRowOff = this.From.RowOff;
        }
        ExcelWorksheet ws = this._drawings.Worksheet;
        double pixOff = pixels - ((ws.GetRowHeight(fromRow + 1) / 0.75) - (fromRowOff / (double)EMU_PER_PIXEL));
        double prevPixOff = pixels;
        int row = fromRow + 1;

        while (pixOff >= 0)
        {
            prevPixOff = pixOff;
            pixOff -= ws.GetRowHeight(++row) / 0.75;
        }
        toRow = row - 1;
        if (fromRow == toRow)
        {
            rowOff = (int)(fromRowOff + (pixels * EMU_PER_PIXEL));
        }
        else
        {
            rowOff = (int)(prevPixOff * EMU_PER_PIXEL);
        }
    }

    internal void SetPixelWidth(double pixels)
    {
        if (this.CellAnchor == eEditAs.TwoCell)
        {
            this._doNotAdjust = true;
            this.GetToColumnFromPixels(pixels, out int col, out int pixOff);

            this.To.Column = col - 2;
            this.To.ColumnOff = pixOff * EMU_PER_PIXEL;
            this._doNotAdjust = false;
        }
        else
        {
            this.Size.Width = (int)Math.Round(pixels * EMU_PER_PIXEL);
        }
    }

    internal void GetToColumnFromPixels(double pixels, out int col, out int colOff, int fromColumn = -1, int fromColumnOff = -1)
    {
        ExcelWorksheet ws = this._drawings.Worksheet;
        decimal mdw = ws.Workbook.MaxFontWidth;
        if(fromColumn<0)
        {
            fromColumn = this.From.Column;
            fromColumnOff = this.From.ColumnOff;
        }
        double pixOff = pixels - (double)(decimal.Truncate(((256 * ws.GetColumnWidth(fromColumn + 1)) + decimal.Truncate(128 / (decimal)mdw)) / 256 * mdw) - (fromColumnOff / EMU_PER_PIXEL));
        double offset = ((double)fromColumnOff / EMU_PER_PIXEL) + pixels;
        col = fromColumn + 2;
        while (pixOff >= 0)
        {
            offset = pixOff;
            pixOff -= (double)decimal.Truncate(((256 * ws.GetColumnWidth(col++)) + decimal.Truncate(128 / (decimal)mdw)) / 256 * mdw);
        }
        colOff = (int)offset;
    }
    #endregion
    #region "Public sizing functions"
    /// <summary>
    /// Set the top left corner of a drawing. 
    /// Note that resizing columns / rows after using this function will effect the position of the drawing
    /// </summary>
    /// <param name="PixelTop">Top pixel</param>
    /// <param name="PixelLeft">Left pixel</param>
    public void SetPosition(int PixelTop, int PixelLeft)
    {
        this.SetPosition(PixelTop, PixelLeft, true);
    }
    internal void SetPosition(int PixelTop, int PixelLeft, bool adjustChildren)
    {
        this._doNotAdjust = true;
        if (this._width == int.MinValue)
        {
            this._width = this.GetPixelWidth();
            this._height = this.GetPixelHeight();
        }
        if(adjustChildren && this.DrawingType == eDrawingType.GroupShape)
        {
            if(this._left== int.MinValue)
            {
                this._left = this.GetPixelLeft();
                this._top = this.GetPixelTop();
            }
            ExcelGroupShape? grp = (ExcelGroupShape)this;
            foreach(ExcelDrawing? d in grp.Drawings)
            {
                d.SetPosition((int)(d._top + (PixelTop - this._top)), (int)(d._left + (PixelLeft - this._left)));
            }
        }

        this.SetPixelTop(PixelTop);
        this.SetPixelLeft(PixelLeft);

        this.SetPixelWidth(this._width);
        this.SetPixelHeight(this._height);
        this._doNotAdjust = false;
    }
    /// <summary>
    /// How the drawing is anchored to the cells.
    /// This effect how the drawing will be resize
    /// <see cref="ChangeCellAnchor(eEditAs, int, int, int, int)"/>
    /// </summary>
    public eEditAs CellAnchor
    {
        get;
        protected set;
    }
    /// <summary>
    /// This will change the cell anchor type, move and resize the drawing.
    /// </summary>
    /// <param name="type">The cell anchor type to change to</param>
    /// <param name="PixelTop">The topmost pixel</param>
    /// <param name="PixelLeft">The leftmost pixel</param>
    /// <param name="width">The width in pixels</param>
    /// <param name="height">The height in pixels</param>
    public void ChangeCellAnchor(eEditAs type, int PixelTop, int PixelLeft, int width, int height)
    {
        this.ChangeCellAnchorTypeInternal(type);
        this.SetPosition(PixelTop, PixelLeft);
        this.SetSize(width, height);
    }
    /// <summary>
    /// This will change the cell anchor type without modifiying the position and size.
    /// </summary>
    /// <param name="type">The cell anchor type to change to</param>
    public void ChangeCellAnchor(eEditAs type)
    {
        if(this.DrawingType==eDrawingType.Control)
        {
            throw new InvalidOperationException("Controls can't change CellAnchor. Must be TwoCell anchor. Please use EditAs property instead.");
        }

        this.GetPositionSize();
        //Save the positions
        double top = this._top;
        double left = this._left;
        double width = this._width;
        double height = this._height;
        //Change the type
        this.ChangeCellAnchorTypeInternal(type);

        //Set the position and size
        this.SetPixelTop(top);
        this.SetPixelLeft(left);

        this.SetPixelWidth(width);
        this.SetPixelHeight(height);
    }

    private void ChangeCellAnchorTypeInternal(eEditAs type)
    {
        if (type != this.CellAnchor)
        {
            this.CellAnchor = type;
            this.RenameNode(this.TopNode, "xdr", $"{type.ToEnumString()}Anchor");
            this.CleanupPositionXml();
            this.SetPositionProperties(this._drawings, this.TopNode);
            this.CellAnchorChanged();
        }
    }

    internal virtual void CellAnchorChanged()
    {
            
    }

    private void CleanupPositionXml()
    {
        switch(this.CellAnchor)
        {
            case eEditAs.OneCell:
                this.DeleteNode("xdr:to");
                this.DeleteNode("xdr:pos");
                this.CreateNode("xdr:from");
                this.CreateNode("xdr:ext");
                break;
            case eEditAs.Absolute:
                this.DeleteNode("xdr:to");
                this.DeleteNode("xdr:from");
                this.CreateNode("xdr:pos");
                this.CreateNode("xdr:ext");
                break;
            default:
                this.DeleteNode("xdr:pos");
                this.DeleteNode("xdr:ext");
                this.CreateNode("xdr:from");
                this.CreateNode("xdr:to");
                break;
        }

    }

    /// <summary>
    /// Set the top left corner of a drawing. 
    /// Note that resizing columns / rows after using this function will effect the position of the drawing
    /// </summary>
    /// <param name="Row">Start row - 0-based index.</param>
    /// <param name="RowOffsetPixels">Offset in pixels</param>
    /// <param name="Column">Start Column - 0-based index.</param>
    /// <param name="ColumnOffsetPixels">Offset in pixels</param>
    public void SetPosition(int Row, int RowOffsetPixels, int Column, int ColumnOffsetPixels)
    {
        this._doNotAdjust = true;

        if (this._width == int.MinValue)
        {
            this._width = this.GetPixelWidth();
            this._height = this.GetPixelHeight();
        }

        this.From.Row = Row;
        this.From.RowOff = RowOffsetPixels * EMU_PER_PIXEL;
        this.From.Column = Column;
        this.From.ColumnOff = ColumnOffsetPixels * EMU_PER_PIXEL;
        if (this.CellAnchor == eEditAs.TwoCell)
        {
            this._left = this.GetPixelLeft();
            this._top = this.GetPixelTop();
        }

        this.SetPixelWidth(this._width);
        this.SetPixelHeight(this._height);
        this._doNotAdjust = false;
    }
    /// <summary>
    /// Set size in Percent.
    /// Note that resizing columns / rows after using this function will effect the size of the drawing
    /// </summary>
    /// <param name="Percent"></param>
    public virtual void SetSize(int Percent)
    {
        this._doNotAdjust = true;
        if (this._width == int.MinValue)
        {
            this._width = this.GetPixelWidth();
            this._height = this.GetPixelHeight();
        }

        this._width *= (double)Percent / 100;
        this._height *= (double)Percent / 100;

        this.SetPixelWidth(this._width);
        this.SetPixelHeight(this._height);
        this._doNotAdjust = false;
    }
    /// <summary>
    /// Set size in pixels
    /// Note that resizing columns / rows after using this function will effect the size of the drawing
    /// </summary>
    /// <param name="PixelWidth">Width in pixels</param>
    /// <param name="PixelHeight">Height in pixels</param>
    public void SetSize(int PixelWidth, int PixelHeight)
    {
        this._doNotAdjust = true;
        this._width = PixelWidth;
        this._height = PixelHeight;
        this.SetPixelWidth(PixelWidth);
        this.SetPixelHeight(PixelHeight);
        this._doNotAdjust = false;
    }
    #endregion
    /// <summary>
    /// Sends the drawing to the back of any overlapping drawings.
    /// </summary>
    public void SendToBack()
    {
        this._drawings.SendToBack(this);
    }
    /// <summary>
    /// Brings the drawing to the front of any overlapping drawings.
    /// </summary>
    public void BringToFront()
    {
        this._drawings.BringToFront(this);
    }
    /// <summary>
    /// Group the drawing together with a list of other drawings. 
    /// <seealso cref="UnGroup(bool)"/>
    /// <seealso cref="ParentGroup"/>
    /// </summary>
    /// <param name="drawing">The drawings to group</param>
    /// <returns>The group shape</returns>
    public ExcelGroupShape Group(params ExcelDrawing[] drawing)
    {
        ExcelGroupShape grp = this._parent;
        foreach(ExcelDrawing? d in drawing)
        {
            ExcelGroupShape.Validate(d, this._drawings, grp);
            if (d._parent != null)
            {
                grp = d._parent;
            }
        }
        grp ??= this._drawings.AddGroupDrawing();
            
        grp.Drawings.AddDrawing(this);

        foreach (ExcelDrawing? d in drawing)
        {
            grp.Drawings.AddDrawing(d);
        }

        grp.SetPositionAndSizeFromChildren();
        return grp;
    }
    internal XmlElement GetFrmxNode(XmlNode node)
    {
        if(node.LocalName == "AlternateContent")
        {
            node = node.FirstChild.FirstChild;
        }

        if(node.LocalName == "sp" || node.LocalName == "pic" || node.LocalName == "cxnSp")
        {
            return (XmlElement)this.CreateNode(node, "xdr:spPr/a:xfrm");
        }
        else if(node.LocalName == "graphicFrame")
        {
            return (XmlElement)this.CreateNode(node, "xdr:xfrm"); 
        }
        return null;
    }

    /// <summary>
    /// Will ungroup this drawing or the entire group, if this drawing is grouped together with other drawings.
    /// If this drawings is not grouped an InvalidOperationException will be returned.
    /// </summary>
    /// <param name="ungroupThisItemOnly">If true this drawing will be removed from the group. 
    /// If it is false, the whole group will be disbanded. If true only this drawing will be removed.
    /// </param>
    public void UnGroup(bool ungroupThisItemOnly=true)
    {
        if(this._parent==null)
        {
            throw new InvalidOperationException("Cannot ungroup this drawing. This drawing is not part of a group");
        }
        if(ungroupThisItemOnly)
        {
            this._parent.Drawings.Remove(this);
        }
        else
        {
            this._parent.Drawings.Clear();
        }           
    }
    /// <summary>
    /// If the drawing is grouped this property contains the Group drawing containing the group.
    /// Otherwise this property is null
    /// </summary>
    public ExcelGroupShape ParentGroup
    { 
        get
        {
            return this._parent;
        }
    }
    internal virtual void DeleteMe()
    {
        this.TopNode.ParentNode.RemoveChild(this.TopNode);            
    }

    /// <summary>
    /// Dispose the object
    /// </summary>
    public virtual void Dispose()
    {
        //TopNode = null;
    }
    internal void GetPositionSize()
    {
        if (this._doNotAdjust)
        {
            return;
        }

        this._top = this.GetPixelTop();
        this._left = this.GetPixelLeft();
        this._height = this.GetPixelHeight();
        this._width = this.GetPixelWidth();
    }
    /// <summary>
    /// Will adjust the position and size of the drawing according to changes in font of rows and to the Normal style.
    /// This method will be called before save, so use it only if you need the coordinates of the drawing.
    /// </summary>
    public void AdjustPositionAndSize()
    {
        if (this._drawings.Worksheet.Workbook._package.DoAdjustDrawings == false)
        {
            return;
        }

        this._drawings.Worksheet.Workbook._package.DoAdjustDrawings = false;
        if (this.EditAs==eEditAs.Absolute)
        {
            this.SetPixelLeft(this._left);
            this.SetPixelTop(this._top);
        }
        if(this.EditAs == eEditAs.Absolute || this.EditAs == eEditAs.OneCell)
        {
            this.SetPixelHeight(this._height);
            this.SetPixelWidth(this._width);
        }

        this._drawings.Worksheet.Workbook._package.DoAdjustDrawings = true;
    }
    internal void UpdatePositionAndSizeXml()
    {
        this.From?.UpdateXml();
        this.To?.UpdateXml();
        this.Size?.UpdateXml();
        this.Position?.UpdateXml();
    }


    internal XmlElement CreateShapeNode()
    {
        XmlElement shapeNode = this.TopNode.OwnerDocument.CreateElement("xdr", "sp", ExcelPackage.schemaSheetDrawings);
        shapeNode.SetAttribute("macro", "");
        shapeNode.SetAttribute("textlink", "");
        this.TopNode.AppendChild(shapeNode);
        return shapeNode;
    }
    internal XmlElement CreateClientData()
    {
        XmlElement clientDataNode = this.TopNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings);
        clientDataNode.SetAttribute("fPrintsWithSheet", "0");
        this.TopNode.GetChildAtPosition(2).GetChildAtPosition(0).GetChildAtPosition(0).AppendChild(clientDataNode);
        this.TopNode.AppendChild(clientDataNode);
        return clientDataNode;
    }
}