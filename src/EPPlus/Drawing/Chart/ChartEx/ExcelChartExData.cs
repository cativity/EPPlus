﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/

using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// The data used as source for the chart. Only spreadsheet internal data is supported at this point.
/// </summary>
public abstract class ExcelChartExData : XmlHelper
{
    string _worksheetName;

    internal ExcelChartExData(string worksheetName, XmlNamespaceManager nsm, XmlNode topNode)
        : base(nsm, topNode) =>
        this._worksheetName = worksheetName;

    /// <summary>
    /// Data formula
    /// </summary>
    public string Formula
    {
        get => this.GetXmlNodeString("cx:f");
        set
        {
            if (ExcelCellBase.IsValidAddress(value))
            {
                this.SetXmlNodeString("cx:f", ExcelCellBase.GetFullAddress(this._worksheetName, value));
            }
            else
            {
                this.SetXmlNodeString("cx:f", value);
            }
        }
    }

    /// <summary>
    /// The direction of the formula
    /// </summary>
    public eFormulaDirection FormulaDirection
    {
        get => this.GetXmlNodeString("cx:f/@dir").ToEnum(eFormulaDirection.Column);
        set => this.SetXmlNodeString("cx:f/@dir", value.ToEnumString());
    }

    /// <summary>
    /// The dimensions name formula. Return null if the element does not exist
    /// </summary>
    public string NameFormula
    {
        get
        {
            if (this.ExistsNode("cx:nf"))
            {
                return this.GetXmlNodeString("cx:nf");
            }
            else
            {
                return null;
            }
        }
        set
        {
            if (ExcelCellBase.IsValidAddress(value))
            {
                this.SetXmlNodeString("cx:nf", ExcelCellBase.GetFullAddress(this._worksheetName, value), true);
            }
            else
            {
                this.SetXmlNodeString("cx:nf", value, true);
            }
        }
    }

    /// <summary>
    /// Direction for the name formula
    /// </summary>
    public eFormulaDirection? NameFormulaDirection
    {
        get
        {
            if (this.ExistsNode("cx:nf"))
            {
                return this.GetXmlNodeString("cx:nf/@dir").ToEnum<eFormulaDirection>(eFormulaDirection.Column);
            }
            else
            {
                return null;
            }
        }
        set
        {
            if (value == null)
            {
                this.DeleteNode("cx:nf/@dir");
            }
            else
            {
                this.SetXmlNodeString("cx:nf/@dir", value.ToEnumString());
            }
        }
    }
}