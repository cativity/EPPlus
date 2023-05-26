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
/// ExcelConditionalFormattingNotContainsErrors
/// </summary>
public class ExcelConditionalFormattingNotContainsErrors
    : ExcelConditionalFormattingRule,
      IExcelConditionalFormattingNotContainsErrors
{
    /****************************************************************************************/

    #region Constructors
    /// <summary>
    /// 
    /// </summary>
    /// <param name="address"></param>
    /// <param name="priority"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    /// <param name="namespaceManager"></param>
    internal ExcelConditionalFormattingNotContainsErrors(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet,
        XmlNode itemElementNode,
        XmlNamespaceManager namespaceManager)
        : base(
               eExcelConditionalFormattingRuleType.NotContainsErrors,
               address,
               priority,
               worksheet,
               itemElementNode,
               (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
    {
        if (itemElementNode==null) //Set default values and create attributes if needed
        {
            this.Formula = string.Format(
                                         "NOT(ISERROR({0}))",
                                         this.Address.Start.Address);
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    internal ExcelConditionalFormattingNotContainsErrors(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet,
        XmlNode itemElementNode)
        : this(
               address,
               priority,
               worksheet,
               itemElementNode,
               null)
    {
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    internal ExcelConditionalFormattingNotContainsErrors(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet)
        : this(
               address,
               priority,
               worksheet,
               null,
               null)
    {
    }
    #endregion Constructors

    /****************************************************************************************/
}