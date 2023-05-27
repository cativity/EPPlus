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
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A collection of datapoints
/// </summary>
public class ExcelChartDataPointCollection : XmlHelper, IEnumerable<ExcelChartDataPoint>
{
    ExcelChart _chart;
    private readonly SortedDictionary<int,ExcelChartDataPoint> _dic = new SortedDictionary<int, ExcelChartDataPoint>();
    internal ExcelChartDataPointCollection(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) : base(ns, topNode)
    {
        this.SchemaNodeOrder = schemaNodeOrder;
        foreach (XmlNode pointNode in this.TopNode.SelectNodes(ExcelChartDataPoint.topNodePath, ns))
        {
            ExcelChartDataPoint? item = new ExcelChartDataPoint(chart, ns, pointNode);
            this._dic.Add(item.Index, item); 
        }

        this._chart = chart;
    }
        
    /// <summary>
    /// Checkes if the index exists in the collection
    /// </summary>
    /// <param name="index">The index</param>
    /// <returns>true if exists</returns>
    public bool ContainsKey(int index)
    {
        return this._dic.ContainsKey(index);
    }
    /// <summary>
    /// Adds a new datapoint to the collection
    /// </summary>
    /// <param name="index">The index</param>
    /// <returns>The datapoint</returns>
    public ExcelChartDataPoint Add(int index)
    {
        return this.AddDp(index, null);
    }

    internal ExcelChartDataPoint AddDp(int idx, string uniqueId=null)
    {
        if (this._dic.ContainsKey(idx))
        {
            throw new ArgumentException($"Point with index {idx} already exists");
        }
        int pos = this.GetItemBefore(idx);

        XmlElement element = this.CreateElement(pos, uniqueId);
        ExcelChartDataPoint? dp = new ExcelChartDataPoint(this._chart, this.NameSpaceManager, element, idx);

        this._dic.Add(idx, dp);

        return dp;
    }

    private XmlElement CreateElement(int idx, string uniqueId="")
    {
        XmlElement pointElement;
        if (this._dic.Count==0)
        {
            pointElement = (XmlElement)this.CreateNode(ExcelChartDataPoint.topNodePath);
        }
        else
        {
            pointElement = this.TopNode.OwnerDocument.CreateElement("c", "dPt", ExcelPackage.schemaChart);
            if(this._dic.ContainsKey(idx))
            {
                this._dic[idx].TopNode.ParentNode.InsertAfter(pointElement, this._dic[idx].TopNode);
            }
            else
            {
                XmlNode? first = this._dic.Values.First().TopNode;
                first.ParentNode.InsertBefore(pointElement,first);
            }
        }
        if(!string.IsNullOrEmpty(uniqueId))
        {
            if(this._chart.IsType3D())
            {
                pointElement.InnerXml = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d contourW=\"25400\"><a:contourClr><a:schemeClr val=\"lt1\"/></a:contourClr></a:sp3d></c:spPr><c:extLst><c:ext xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" uri = \"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"><c16:uniqueId val=\"{" + uniqueId + "}\"/></c:ext></c:extLst>";
            }
            else
            {
                pointElement.InnerXml = "<c:extLst><c:ext xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" uri = \"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"><c16:uniqueId val=\"{" + uniqueId + "}\"/></c:ext></c:extLst>";
            }
        }
        return pointElement;
    }

    private int GetItemBefore(int index)
    {
        if(this._dic.ContainsKey(index-1))
        {
            return index-1;
        }
        int retIx=-1;
        foreach (int ix in this._dic.Keys.OrderBy(x=>x))
        {
            if(index < ix)
            {
                return retIx;
            }
            retIx = ix;
        }
        return retIx;
    }
    /// <summary>
    /// Indexer
    /// </summary>
    /// <param name="index">The index</param>
    /// <returns></returns>
    public ExcelChartDataPoint this[int index]
    {
        get
        {
            return this._dic[index];
        }
    }
    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get
        {
            return this._dic.Count;
        }
    }
    /// <summary>
    /// Gets the enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelChartDataPoint> GetEnumerator()
    {
        return this._dic.Values.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._dic.Values.GetEnumerator();
    }
}