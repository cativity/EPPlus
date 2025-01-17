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
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A date group
/// </summary>
public class ExcelPivotTableFieldDateGroup : ExcelPivotTableFieldGroup
{
    internal ExcelPivotTableFieldDateGroup(XmlNamespaceManager ns, XmlNode topNode)
        : base(ns, topNode)
    {
    }

    const string groupByPath = "d:fieldGroup/d:rangePr/@groupBy";

    /// <summary>
    /// How to group the date field
    /// </summary>
    public eDateGroupBy GroupBy
    {
        get
        {
            string v = this.GetXmlNodeString(groupByPath);

            if (v != "")
            {
                return (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), v, true);
            }
            else
            {
                throw new Exception("Invalid date Groupby");
            }
        }
        private set => this.SetXmlNodeString(groupByPath, value.ToString().ToLower(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Auto detect start date
    /// </summary>
    public bool AutoStart => this.GetXmlNodeBool("@autoStart", false);

    /// <summary>
    /// Auto detect end date
    /// </summary>
    public bool AutoEnd => this.GetXmlNodeBool("@autoStart", false);
}