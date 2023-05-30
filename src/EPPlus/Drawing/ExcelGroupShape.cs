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

using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// A collection of child drawings to a group drawing
/// </summary>
public class ExcelDrawingsGroup : IEnumerable<ExcelDrawing>, IDisposable
{
    private ExcelGroupShape _parent;
    internal Dictionary<string, int> _drawingNames;
    private List<ExcelDrawing> _groupDrawings;
    XmlNamespaceManager _nsm;
    XmlNode _topNode;

    internal ExcelDrawingsGroup(ExcelGroupShape parent, XmlNamespaceManager nsm, XmlNode topNode)
    {
        this._parent = parent;
        this._nsm = nsm;
        this._topNode = topNode;
        this._drawingNames = new Dictionary<string, int>();
        this.AddDrawings();
    }

    private void AddDrawings()
    {
        this._groupDrawings = new List<ExcelDrawing>();

        foreach (XmlNode node in this._topNode.ChildNodes)
        {
            if (node.LocalName != "nvGrpSpPr" && node.LocalName != "grpSpPr")
            {
                ExcelDrawing? grpDraw = ExcelDrawing.GetDrawingFromNode(this._parent._drawings, node, (XmlElement)node, this._parent);
                this._groupDrawings.Add(grpDraw);
                this._drawingNames.Add(grpDraw.Name, this._groupDrawings.Count - 1);
            }
        }
    }

    /// <summary>
    /// Adds a drawing to the group
    /// </summary>
    /// <param name="drawing"></param>
    public void Add(ExcelDrawing drawing)
    {
        this.CheckNotDisposed();
        this.AddDrawing(drawing);
        drawing.ParentGroup.SetPositionAndSizeFromChildren();
    }

    private void CheckNotDisposed()
    {
        if (this._topNode == null)
        {
            throw new ObjectDisposedException("This group drawing has been disposed.");
        }
    }

    internal void AddDrawing(ExcelDrawing drawing)
    {
        if (drawing._parent == this._parent)
        {
            return; //This drawing is already added to the group, exit
        }

        ExcelGroupShape.Validate(drawing, drawing._drawings, this._parent);
        this.AdjustXmlAndMoveToGroup(drawing);
        ExcelGroupShape.Validate(drawing, this._parent._drawings, this._parent);
        this.AppendDrawingNode(drawing.TopNode);
        drawing._parent = this._parent;

        this._groupDrawings.Add(drawing);
        this._drawingNames.Add(drawing.Name, this._groupDrawings.Count - 1);
    }

    private void AdjustXmlAndMoveToGroup(ExcelDrawing d)
    {
        d._drawings.RemoveDrawing(d._drawings._drawingsList.IndexOf(d), false);
        double height = d.GetPixelHeight();
        double width = d.GetPixelWidth();
        int top = d.GetPixelTop();
        int left = d.GetPixelLeft();
        XmlNode? node = d.TopNode.GetChildAtPosition(2);
        XmlElement xFrmNode = d.GetFrmxNode(node);

        if (xFrmNode.ChildNodes.Count == 0)
        {
            _ = d.CreateNode(xFrmNode, "a:off");
            _ = d.CreateNode(xFrmNode, "a:ext");
        }

        XmlElement? offNode = (XmlElement)xFrmNode.SelectSingleNode("a:off", this._nsm);
        offNode.SetAttribute("y", (top * ExcelDrawing.EMU_PER_PIXEL).ToString());
        offNode.SetAttribute("x", (left * ExcelDrawing.EMU_PER_PIXEL).ToString());
        XmlElement? extNode = (XmlElement)xFrmNode.SelectSingleNode("a:ext", this._nsm);
        extNode.SetAttribute("cy", Math.Round(height * ExcelDrawing.EMU_PER_PIXEL, 0).ToString());
        extNode.SetAttribute("cx", Math.Round(width * ExcelDrawing.EMU_PER_PIXEL, 0).ToString());

        d.SetGroupChild(offNode, extNode);
        _ = node.ParentNode.RemoveChild(node);

        if (d.TopNode.ParentNode?.ParentNode?.LocalName == "AlternateContent")
        {
            XmlNode? containerNode = d.TopNode.ParentNode?.ParentNode;
            _ = d.TopNode.ParentNode.RemoveChild(d.TopNode);
            _ = containerNode.ParentNode.RemoveChild(containerNode);
            _ = containerNode.FirstChild.AppendChild(node);
            node = containerNode;
        }
        else
        {
            _ = d.TopNode.ParentNode.RemoveChild(d.TopNode);
        }

        d.AdjustXPathsForGrouping(true);
        d.TopNode = node;
    }

    private void AdjustXmlAndMoveFromGroup(ExcelDrawing d)
    {
        double height = d.GetPixelHeight();
        double width = d.GetPixelWidth();
        int top = d.GetPixelTop();
        int left = d.GetPixelLeft();
        XmlNode drawingNode;

        if (this._parent.TopNode.ParentNode?.ParentNode?.LocalName == "AlternateContent") //Create alternat content above ungrouped drawing.
        {
            //drawingNode = xmlDoc.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            drawingNode = this._parent.TopNode.ParentNode.ParentNode.CloneNode(false);
            XmlNode? choiceNode = this._parent.TopNode.ParentNode.CloneNode(false);
            _ = drawingNode.AppendChild(choiceNode);
            _ = d.TopNode.ParentNode.RemoveChild(d.TopNode);
            _ = choiceNode.AppendChild(d.TopNode);
            drawingNode = this.CreateAnchorNode(drawingNode);
            XmlNode? addBeforeNode = this._parent.TopNode.ParentNode.ParentNode;
            _ = addBeforeNode.ParentNode.InsertBefore(drawingNode, addBeforeNode);
        }
        else
        {
            _ = d.TopNode.ParentNode.RemoveChild(d.TopNode);
            drawingNode = this.CreateAnchorNode(d.TopNode);
            _ = this._parent.TopNode.ParentNode.InsertBefore(drawingNode, this._parent.TopNode);
        }

        d.AdjustXPathsForGrouping(false);
        d.TopNode = drawingNode;
        d.SetPosition(top, left);
        d.SetSize((int)width, (int)height);
    }

    private XmlNode CreateAnchorNode(XmlNode drawingNode)
    {
        XmlNode? topNode = this._parent.TopNode.CloneNode(false);
        _ = topNode.AppendChild(this._parent.TopNode.GetChildAtPosition(0).CloneNode(true));
        _ = topNode.AppendChild(this._parent.TopNode.GetChildAtPosition(1).CloneNode(true));
        _ = topNode.AppendChild(drawingNode);
        int ix = 3;

        while (ix < this._parent.TopNode.ChildNodes.Count)
        {
            _ = topNode.AppendChild(this._parent.TopNode.ChildNodes[ix].CloneNode(true));
            ix++;
        }

        return topNode;
    }

    private void AppendDrawingNode(XmlNode drawingNode)
    {
        if (drawingNode.ParentNode?.ParentNode?.LocalName == "AlternateContent")
        {
            _ = this._topNode.AppendChild(drawingNode.ParentNode.ParentNode);
        }
        else
        {
            _ = this._topNode.AppendChild(drawingNode);
        }
    }

    /// <summary>
    /// Disposes the class
    /// </summary>
    public void Dispose()
    {
        this._parent = null;
        this._topNode = null;
    }

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get { return this._groupDrawings.Count; }
    }

    /// <summary>
    /// Returns the drawing at the specified position.  
    /// </summary>
    /// <param name="PositionID">The position of the drawing. 0-base</param>
    /// <returns></returns>
    public ExcelDrawing this[int PositionID]
    {
        get { return this._groupDrawings[PositionID]; }
    }

    /// <summary>
    /// Returns the drawing matching the specified name
    /// </summary>
    /// <param name="Name">The name of the worksheet</param>
    /// <returns></returns>
    public ExcelDrawing this[string Name]
    {
        get
        {
            if (this._drawingNames.ContainsKey(Name))
            {
                return this._groupDrawings[this._parent._drawings._drawingNames[Name]];
            }
            else
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Gets the enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelDrawing> GetEnumerator()
    {
        return this._groupDrawings.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._groupDrawings.GetEnumerator();
    }

    /// <summary>
    /// Removes the <see cref="ExcelDrawing"/> from the group
    /// </summary>
    /// <param name="drawing">The drawing to remove</param>
    public void Remove(ExcelDrawing drawing)
    {
        this.CheckNotDisposed();
        _ = this._groupDrawings.Remove(drawing);
        this.AdjustXmlAndMoveFromGroup(drawing);
        int ix = this._parent._drawings._drawingsList.IndexOf(this._parent);
        this._parent._drawings._drawingsList.Insert(ix, drawing);

        //Remove 
        if (this._parent.Drawings.Count == 0)
        {
            this._parent._drawings.Remove(this._parent);
        }

        this._parent._drawings.ReIndexNames(ix, 1);
        drawing._parent = null;
    }

    /// <summary>
    /// Removes all children drawings from the group.
    /// </summary>
    public void Clear()
    {
        this.CheckNotDisposed();

        while (this._groupDrawings.Count > 0)
        {
            this.Remove(this._groupDrawings[0]);
        }
    }
}

/// <summary>
/// Grouped shapes
/// </summary>
public class ExcelGroupShape : ExcelDrawing
{
    internal ExcelGroupShape(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null)
        : base(drawings, node, "xdr:grpSp", "xdr:nvGrpSpPr/xdr:cNvPr", parent)
    {
        XmlNode? grpNode = this.CreateNode(this._topPath);

        if (grpNode.InnerXml == "")
        {
            grpNode.InnerXml =
                "<xdr:nvGrpSpPr><xdr:cNvPr name=\"\" id=\"3\"><a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId id=\"{F33F4CE3-706D-4DC2-82DA-B596E3C8ACD0}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off y=\"0\" x=\"0\"/><a:ext cy=\"0\" cx=\"0\"/><a:chOff y=\"0\" x=\"0\"/><a:chExt cy=\"0\" cx=\"0\"/></a:xfrm></xdr:grpSpPr>";
        }

        if (parent == null)
        {
            _ = this.CreateNode("xdr:clientData");
        }
    }

    ExcelDrawingsGroup _groupDrawings;

    /// <summary>
    /// A collection of shapes
    /// </summary>
    public ExcelDrawingsGroup Drawings
    {
        get
        {
            if (this._groupDrawings == null)
            {
                if (string.IsNullOrEmpty(this._topPath))
                {
                    this._groupDrawings = new ExcelDrawingsGroup(this, this.NameSpaceManager, this.TopNode);
                }
                else
                {
                    this._groupDrawings = new ExcelDrawingsGroup(this, this.NameSpaceManager, this.GetNode(this._topPath));
                }
            }

            return this._groupDrawings;
        }
    }

    internal static void Validate(ExcelDrawing d, ExcelDrawings drawings, ExcelGroupShape grp)
    {
        if (d._drawings != drawings)
        {
            throw new InvalidOperationException("All drawings must be in the same worksheet.");
        }

        if (d._parent != null && d._parent != grp)
        {
            throw new InvalidOperationException($"The drawing {d.Name} is already in a group different from the other drawings.");
        }
    }

    internal void SetPositionAndSizeFromChildren()
    {
        ExcelDrawing? pd = this.Drawings[0];
        pd.GetPositionSize();

        double t = pd._top,
               l = pd._left,
               b = pd._top + pd._height,
               r = pd._left + pd._width;

        for (int i = 1; i < this.Drawings.Count; i++)
        {
            ExcelDrawing? d = this.Drawings[i];
            d.GetPositionSize();

            if (t > d._top)
            {
                t = d._top;
            }

            if (l > d._left)
            {
                l = d._left;
            }

            if (r < d._left + d._width)
            {
                r = d._left + d._width;
            }

            if (b < d._top + d._height)
            {
                b = d._top + d._height;
            }
        }

        this.SetPosition((int)t, (int)l, false);
        this.SetSize((int)(r - l), (int)(b - t));

        this.SetxFrmPosition();
    }

    internal void AdjustChildrenForResizeRow(double prevTop)
    {
        int top = this.GetPixelTop();
        double diff = top - prevTop;

        if (diff != 0)
        {
            for (int i = 0; i < this.Drawings.Count; i++)
            {
                this.Drawings[i].SetPixelTop(this.Drawings[i]._top + diff);
                this.Drawings[i].Position.UpdateXml();
            }
        }
    }

    internal void AdjustChildrenForResizeColumn(double prevLeft)
    {
        int left = this.GetPixelLeft();
        double diff = left - prevLeft;

        if (diff != 0)
        {
            for (int i = 0; i < this.Drawings.Count; i++)
            {
                this.Drawings[i].SetPixelLeft(this.Drawings[i]._left + diff);
                this.Drawings[i].Position.UpdateXml();
            }
        }
    }

    private void SetxFrmPosition()
    {
        this.xFrmPosition.X = (int)(this._left * EMU_PER_PIXEL);
        this.xFrmPosition.Y = (int)(this._top * EMU_PER_PIXEL);
        this.xFrmSize.Width = (long)(this._width * EMU_PER_PIXEL);
        this.xFrmSize.Height = (long)(this._height * EMU_PER_PIXEL);

        this.xFrmChildPosition.X = (int)(this._left * EMU_PER_PIXEL);
        this.xFrmChildPosition.Y = (int)(this._top * EMU_PER_PIXEL);
        this.xFrmChildSize.Width = (long)(this._width * EMU_PER_PIXEL);
        this.xFrmChildSize.Height = (long)(this._height * EMU_PER_PIXEL);
    }

    ExcelDrawingCoordinate _xFrmPosition;

    internal ExcelDrawingCoordinate xFrmPosition
    {
        get { return this._xFrmPosition ??= new ExcelDrawingCoordinate(this.NameSpaceManager, this.GetNode("xdr:grpSp/xdr:grpSpPr/a:xfrm/a:off")); }
    }

    ExcelDrawingSize _xFrmSize;

    internal ExcelDrawingSize xFrmSize
    {
        get { return this._xFrmSize ??= new ExcelDrawingSize(this.NameSpaceManager, this.GetNode("xdr:grpSp/xdr:grpSpPr/a:xfrm/a:ext")); }
    }

    ExcelDrawingCoordinate _xFrmChildPosition;

    internal ExcelDrawingCoordinate xFrmChildPosition
    {
        get { return this._xFrmChildPosition ??= new ExcelDrawingCoordinate(this.NameSpaceManager, this.GetNode("xdr:grpSp/xdr:grpSpPr/a:xfrm/a:chOff")); }
    }

    ExcelDrawingSize _xFrmChildSize;

    internal ExcelDrawingSize xFrmChildSize
    {
        get { return this._xFrmChildSize ??= new ExcelDrawingSize(this.NameSpaceManager, this.GetNode("xdr:grpSp/xdr:grpSpPr/a:xfrm/a:chExt")); }
    }

    /// <summary>
    /// The type of drawing
    /// </summary>
    public override eDrawingType DrawingType
    {
        get { return eDrawingType.GroupShape; }
    }
}