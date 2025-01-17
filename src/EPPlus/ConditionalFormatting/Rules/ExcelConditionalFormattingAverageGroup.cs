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
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

/// <summary>
/// ExcelConditionalFormattingAverageGroup
/// </summary>
public class ExcelConditionalFormattingAverageGroup : ExcelConditionalFormattingRule, IExcelConditionalFormattingAverageGroup
{
    /****************************************************************************************/

    #region Constructors

    /// <summary>
    /// 
    /// </summary>
    /// <param name="type"></param>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    /// <param name="namespaceManager"></param>
    internal ExcelConditionalFormattingAverageGroup(eExcelConditionalFormattingRuleType type,
                                                    ExcelAddress address,
                                                    int priority,
                                                    ExcelWorksheet worksheet,
                                                    XmlNode itemElementNode,
                                                    XmlNamespaceManager namespaceManager)
        : base(type, address, priority, worksheet, itemElementNode, namespaceManager == null ? worksheet.NameSpaceManager : namespaceManager)
    {
    }

    /// <summary>
    /// 
    /// </summary>
    ///<param name="type"></param>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    internal ExcelConditionalFormattingAverageGroup(eExcelConditionalFormattingRuleType type,
                                                    ExcelAddress address,
                                                    int priority,
                                                    ExcelWorksheet worksheet,
                                                    XmlNode itemElementNode)
        : this(type, address, priority, worksheet, itemElementNode, null)
    {
    }

    /// <summary>
    /// 
    /// </summary>
    ///<param name="type"></param>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    internal ExcelConditionalFormattingAverageGroup(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
        : this(type, address, priority, worksheet, null, null)
    {
    }

    #endregion Constructors

    /****************************************************************************************/
}