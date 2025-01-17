﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/

using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// A collection of individual data labels
/// </summary>
public class ExcelChartExDataLabelCollection : ExcelChartExDataLabel, IDrawingStyle, IEnumerable<ExcelChartExDataLabelItem>
{
    SortedDictionary<int, ExcelChartExDataLabelItem> _dic = new SortedDictionary<int, ExcelChartExDataLabelItem>();

    internal ExcelChartExDataLabelCollection(ExcelChartExSerie serie, XmlNamespaceManager ns, XmlNode node, string[] schemaNodeOrder)
        : base(serie, ns, node)
    {
        this._chart = serie._chart;
        this.AddSchemaNodeOrder(schemaNodeOrder, new string[] { "numFmt", "spPr", "txPr", "visibility", "separator" });

        foreach (XmlNode pointNode in this.TopNode.SelectNodes(_dataLabelPath, ns))
        {
            ExcelChartExDataLabelItem? item = new ExcelChartExDataLabelItem(serie, ns, pointNode);
            this._dic.Add(item.Index, item);
        }
    }

    void IDrawingStyleBase.CreatespPr() => this.CreatespPrNode("cx:spPr");

    /// <summary>
    /// Adds an individual data label for customization.
    /// </summary>
    /// <param name="index">The zero based index</param>
    /// <returns></returns>
    public ExcelChartExDataLabelItem Add(int index)
    {
        if (this._dic.ContainsKey(index))
        {
            throw new InvalidOperationException($"Data label with index {index} already exists.");
        }

        XmlNode? node = this._serie.CreateNode("cx:dataLabels/cx:dataLabel", false, true);

        return new ExcelChartExDataLabelItem(this._serie, this.NameSpaceManager, node, index);
    }

    /// <summary>
    /// Returns tje data label at the specific position.  
    /// </summary>
    /// <param name="index">The index of the datalabel. 0-base.</param>
    /// <returns>Returns null if the data label does not exist in the collection</returns>
    public ExcelChartExDataLabel this[int index]
    {
        get
        {
            if (this._dic.ContainsKey(index))
            {
                return this._dic[index];
            }

            return null;
        }
    }

    /// <summary>
    /// Get the enumerator
    /// </summary>
    /// <returns></returns>
    public IEnumerator<ExcelChartExDataLabelItem> GetEnumerator() => this._dic.Values.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this._dic.Values.GetEnumerator();
}