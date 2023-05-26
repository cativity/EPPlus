﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/29/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// String data reference for an extended chart
/// </summary>
public class ExcelChartExStringData : ExcelChartExData
{
    internal ExcelChartExStringData(string worksheetName, XmlNamespaceManager nsm, XmlNode topNode) : base(worksheetName, nsm, topNode)
    {
    }
    /// <summary>
    /// The type of data
    /// </summary>
    public eStringDataType Type 
    {
        get
        {
            string? s = this.GetXmlNodeString("@type");
            switch (s)
            {
                case "entityId":
                    return eStringDataType.EntityId;
                case "colorStr":
                    return eStringDataType.ColorString;
                default:
                    return eStringDataType.Category;
            }
        }
        set
        {
            string s;
            switch (value)
            {
                case eStringDataType.EntityId:
                    s = "entityId";
                    break;
                case eStringDataType.ColorString:
                    s = "colorStr";
                    break;
                default:
                    s = "cat";
                    break;
            }

            this.SetXmlNodeString("@type", s);
        }
    }
}