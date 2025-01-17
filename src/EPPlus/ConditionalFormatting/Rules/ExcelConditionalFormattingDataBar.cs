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
using System.Globalization;

namespace OfficeOpenXml.ConditionalFormatting;

/// <summary>
/// Databar
/// </summary>
public class ExcelConditionalFormattingDataBar : ExcelConditionalFormattingRule, IExcelConditionalFormattingDataBarGroup
{
    /****************************************************************************************/

    #region Private Properties

    #endregion Private Properties

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
    internal ExcelConditionalFormattingDataBar(eExcelConditionalFormattingRuleType type,
                                               ExcelAddress address,
                                               int priority,
                                               ExcelWorksheet worksheet,
                                               XmlNode itemElementNode,
                                               XmlNamespaceManager namespaceManager)
        : base(type, address, priority, worksheet, itemElementNode, namespaceManager == null ? worksheet.NameSpaceManager : namespaceManager)
    {
        string[]? s = this.SchemaNodeOrder;
        Array.Resize(ref s, s.Length + 2); //Fixes issue 15429. Append node order instead om overwriting it.
        s[s.Length - 2] = "cfvo";
        s[s.Length - 1] = "color";
        this.SchemaNodeOrder = s;

        //Create the <dataBar> node inside the <cfRule> node
        if (itemElementNode != null && itemElementNode.HasChildNodes)
        {
            bool high = false;

            foreach (XmlNode node in itemElementNode.SelectNodes("d:dataBar/d:cfvo", this.NameSpaceManager))
            {
                if (high == false)
                {
                    this.LowValue = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, node, namespaceManager);
                    high = true;
                }
                else
                {
                    this.HighValue = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, node, namespaceManager);
                }
            }
        }
        else
        {
            XmlNode? iconSetNode = this.CreateComplexNode(this.Node, ExcelConditionalFormattingConstants.Paths.DataBar);

            XmlElement? lowNode = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
            _ = iconSetNode.AppendChild(lowNode);

            this.LowValue = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Min,
                                                                           0,
                                                                           "",
                                                                           eExcelConditionalFormattingRuleType.DataBar,
                                                                           address,
                                                                           priority,
                                                                           worksheet,
                                                                           lowNode,
                                                                           namespaceManager);

            XmlElement? highNode = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
            _ = iconSetNode.AppendChild(highNode);

            this.HighValue = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Max,
                                                                            0,
                                                                            "",
                                                                            eExcelConditionalFormattingRuleType.DataBar,
                                                                            address,
                                                                            priority,
                                                                            worksheet,
                                                                            highNode,
                                                                            namespaceManager);
        }

        this.Type = type;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="type"></param>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    internal ExcelConditionalFormattingDataBar(eExcelConditionalFormattingRuleType type,
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
    /// <param name="type"></param>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    internal ExcelConditionalFormattingDataBar(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
        : this(type, address, priority, worksheet, null, null)
    {
    }

    #endregion Constructors

    private const string _showValuePath = "d:dataBar/@showValue";

    /// <summary>
    /// If true the values of the cells are shown, otherwise the cells only contain the databars.
    /// </summary>
    public bool ShowValue
    {
        get => this.GetXmlNodeBool(_showValuePath, true);
        set => this.SetXmlNodeBool(_showValuePath, value);
    }

    /// <summary>
    /// The low value
    /// </summary>
    public ExcelConditionalFormattingIconDataBarValue LowValue { get; internal set; }

    /// <summary>
    /// The high value
    /// </summary>
    public ExcelConditionalFormattingIconDataBarValue HighValue { get; internal set; }

    private const string _colorPath = "d:dataBar/d:color/@rgb";

    /// <summary>
    /// The color of the databar
    /// </summary>
    public Color Color
    {
        get
        {
            string? rgb = this.GetXmlNodeString(_colorPath);

            if (!string.IsNullOrEmpty(rgb))
            {
                return Color.FromArgb(int.Parse(rgb, NumberStyles.HexNumber));
            }

            return Color.White;
        }
        set => this.SetXmlNodeString(_colorPath, value.ToArgb().ToString("X"));
    }
}