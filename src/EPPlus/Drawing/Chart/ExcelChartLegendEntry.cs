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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// An individual serie item within the chart legend
/// </summary>
public class ExcelChartLegendEntry : XmlHelper, IDrawingStyle
{
    internal ExcelChartStandard _chart;
    internal ExcelChartLegendEntry(XmlNamespaceManager nsm, XmlNode topNode, ExcelChartStandard chart) : base(nsm, topNode)
    {
        this.Init(chart);
        this.Index = this.GetXmlNodeInt("c:idx/@val");
        this.HasValue = true;
    }

    internal ExcelChartLegendEntry(XmlNamespaceManager nsm, XmlNode legendNode, ExcelChartStandard chart, int serieIndex) : base(nsm)
    {
        this.Init(chart);
        this.TopNode = legendNode;
        this.Index = serieIndex;
    }
    private void Init(ExcelChartStandard chart)
    {
        this._chart = chart;
        this.SchemaNodeOrder = new string[] { "idx", "deleted", "txPr" };
    }
    /// <summary>
    /// The index of the item
    /// </summary>
    public int Index
    {
        get;
        internal set;

    }
    /// <summary>
    /// If the items has been deleted or is visible.
    /// </summary>
    public bool Deleted
    {
        get
        {
            return this.GetXmlNodeBool("c:delete/@val");
        }
        set
        {
            this.CreateTopNode();
            this.HasValue = true;
            this.SetXmlNodeBool("c:delete/@val", value);
        }
    }
    internal bool HasValue { get; set; }
    private void CreateTopNode()
    {
        if(this.TopNode.LocalName != "legendEntry")
        {
            ExcelChartLegend? legend = this._chart.Legend;
            int preIx = legend.GetPreEntryIndex(this.Index);
            XmlNode legendEntryNode;
            if (preIx == -1)
            {
                legendEntryNode = legend.CreateNode("c:legendEntry", false, true);
            }
            else
            {
                legendEntryNode = this._chart.ChartXml.CreateElement("c", "legendEntry", ExcelPackage.schemaChart);
                XmlNode? refNode = legend.Entries[preIx].TopNode;
                refNode.ParentNode.InsertBefore(legendEntryNode, refNode);
            }

            this.TopNode = legendEntryNode;
            this.SetXmlNodeInt("c:idx/@val", this.Index);
        }
    }

    ExcelTextFont _font = null;
    /// <summary>
    /// The Font properties
    /// </summary>
    public ExcelTextFont Font
    {
        get
        {
            if (this._font == null)
            {
                this.CreateTopNode();
                this._font = new ExcelTextFont(this._chart, this.NameSpaceManager, this.TopNode, $"c:txPr/a:p/a:pPr/a:defRPr", this.SchemaNodeOrder, this.InitChartXml);                    
            }
            return this._font;
        }
    }
    internal void InitChartXml()
    {
        if (this.HasValue)
        {
            return;
        }

        this.HasValue = true;
        this._font.CreateTopNode();
        if (this._chart.StyleManager.Style == null)
        {
            return;
        }

        if (this._chart.StyleManager.Style.Legend.HasTextRun)
        {
            XmlElement? node = (XmlElement)this.CreateNode("c:txPr/a:p/a:pPr/a:defRPr");
            CopyElement(this._chart.StyleManager.Style.Legend.DefaultTextRun.PathElement, node);
        }
        if (this._chart.StyleManager.Style.Legend.HasTextBody)
        {
            XmlElement? node = (XmlElement)this.CreateNode("c:txPr/a:bodyPr");
            CopyElement(this._chart.StyleManager.Style.Legend.DefaultTextBody.PathElement, node);
        }
    }
    ExcelTextBody _textBody = null;
    /// <summary>
    /// Access to text body properties
    /// </summary>
    public ExcelTextBody TextBody
    {
        get
        {
            return this._textBody ??= new ExcelTextBody(this.NameSpaceManager, this.TopNode, $"c:txPr/a:bodyPr", this.SchemaNodeOrder);
        }
    }

    /// <summary>
    /// Access to border properties
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get
        {
            return null;
        }
    }
    /// <summary>
    /// Access to effects styling properties
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            return null;
        }
    }

    /// <summary>
    /// Access to fill styling properties.
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            return null;
        }
    }

    /// <summary>
    /// Access to 3D properties.
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get
        {
            return null;
        }
    }


    internal void Save()
    {
        if(this.Deleted==true)
        {
            this.DeleteNode("c:txPr");
        }
        else
        {
            if (this.ExistsNode("c:txPr"))
            {
                this.DeleteNode("c:delete");
            }
            else
            {
                this.TopNode.ParentNode.RemoveChild(this.TopNode);
            }
        }
    }

    void IDrawingStyleBase.CreatespPr()
    {
            
    }
}