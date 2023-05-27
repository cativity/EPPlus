/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/

using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// An individual data point
/// </summary>
public class ExcelChartExDataPoint : XmlHelper, IDrawingStyleBase
{
    ExcelChartExSerie _serie;

    internal ExcelChartExDataPoint(ExcelChartExSerie serie, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder)
        : base(ns, topNode)
    {
        this._serie = serie;
        this.AddSchemaNodeOrder(schemaNodeOrder, new string[] { "spPr" });
        this.Index = this.GetXmlNodeInt(indexPath);
    }

    internal ExcelChartExDataPoint(ExcelChartExSerie serie, XmlNamespaceManager ns, XmlNode topNode, int index, string[] schemaNodeOrder)
        : base(ns, topNode)
    {
        this._serie = serie;
        this.AddSchemaNodeOrder(schemaNodeOrder, new string[] { "spPr" });
        this.Index = index;
    }

    internal const string dataPtPath = "cx:dataPt";
    internal const string SubTotalPath = "cx:layoutPr/cx:subtotals/cx:idx";
    const string indexPath = "@idx";

    /// <summary>
    /// The index of the datapoint
    /// </summary>
    public int Index { get; private set; }

    /// <summary>
    /// The data point is a subtotal. Applies for waterfall charts.
    /// </summary>
    public bool SubTotal
    {
        get { return this.ExistsNode($"{this.GetSubTotalPath()}[@val={this.Index}]"); }
        set
        {
            string? path = this.GetSubTotalPath();

            if (value)
            {
                if (!this.ExistsNode($"{path}[@val={this.Index}]"))
                {
                    XmlElement? idxElement = (XmlElement)this.CreateNode(path, false, true);
                    idxElement.SetAttribute("val", this.Index.ToString(CultureInfo.InvariantCulture));
                }
            }
            else
            {
                this.DeleteNode($"{path}/[@val={this.Index}]");
            }
        }
    }

    private string GetSubTotalPath()
    {
        if (this.TopNode.LocalName == "series")
        {
            return "cx:layoutPr/cx:subtotals/cx:idx";
        }
        else
        {
            return "../cx:layoutPr/cx:subtotals/cx:idx";
        }
    }

    ExcelDrawingFill _fill = null;

    /// <summary>
    /// A reference to fill properties
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            if (this._fill == null)
            {
                this.CreateDp();
                this._fill = new ExcelDrawingFill(this._serie._chart, this.NameSpaceManager, this.TopNode, "cx:spPr", this.SchemaNodeOrder);
            }

            return this._fill;
        }
    }

    ExcelDrawingBorder _line = null;

    /// <summary>
    /// A reference to line properties
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get
        {
            if (this._line == null)
            {
                this.CreateDp();
                this._line = new ExcelDrawingBorder(this._serie._chart, this.NameSpaceManager, this.TopNode, "cx:spPr/a:ln", this.SchemaNodeOrder);
            }

            return this._line;
        }
    }

    private ExcelDrawingEffectStyle _effect = null;

    /// <summary>
    /// A reference to line properties
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            if (this._effect == null)
            {
                this.CreateDp();

                this._effect = new ExcelDrawingEffectStyle(this._serie._chart,
                                                           this.NameSpaceManager,
                                                           this.TopNode,
                                                           "cx:spPr/a:effectLst",
                                                           this.SchemaNodeOrder);
            }

            return this._effect;
        }
    }

    ExcelDrawing3D _threeD = null;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get
        {
            if (this._threeD == null)
            {
                this.CreateDp();
                this._threeD = new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "cx:spPr", this.SchemaNodeOrder);
            }

            return this._threeD;
        }
    }

    private void CreateDp()
    {
        if (this.TopNode.LocalName == "series")
        {
            XmlElement pointElement;
            XmlElement? prepend = this.GetPrependItem();

            if (prepend == null)
            {
                pointElement = (XmlElement)this.CreateNode(dataPtPath);
            }
            else
            {
                pointElement = this.TopNode.OwnerDocument.CreateElement(dataPtPath, ExcelPackage.schemaChartExMain);
                prepend.ParentNode.InsertBefore(pointElement, prepend);
            }

            pointElement.SetAttribute("idx", this.Index.ToString(CultureInfo.InvariantCulture));
            this.TopNode = pointElement;
        }
    }

    private XmlElement GetPrependItem()
    {
        SortedDictionary<int, ExcelChartExDataPoint>? dic = this._serie.DataPoints._dic;
        int prevKey = -1;

        foreach (ExcelChartExDataPoint? v in dic.Values)
        {
            if (v.TopNode.LocalName == "dataPt" && prevKey < v.Index)
            {
                return (XmlElement)v.TopNode;
            }
        }

        return null;
    }

    void IDrawingStyleBase.CreatespPr()
    {
        this.CreatespPrNode("cx:spPr");
    }
}