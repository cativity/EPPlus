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
/// Conditional formatting with a three icon set
/// </summary>
public class ExcelConditionalFormattingThreeIconSet : ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>
{
    internal ExcelConditionalFormattingThreeIconSet(ExcelAddress address,
                                                    int priority,
                                                    ExcelWorksheet worksheet,
                                                    XmlNode itemElementNode,
                                                    XmlNamespaceManager namespaceManager)
        : base(eExcelConditionalFormattingRuleType.ThreeIconSet,
               address,
               priority,
               worksheet,
               itemElementNode,
               namespaceManager == null ? worksheet.NameSpaceManager : namespaceManager)
    {
    }
}

/// <summary>
/// ExcelConditionalFormattingThreeIconSet
/// </summary>
public class ExcelConditionalFormattingIconSetBase<T> : ExcelConditionalFormattingRule, IExcelConditionalFormattingThreeIconSet<T>
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
    /// <param name="address"></param>
    /// <param name="priority"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    /// <param name="namespaceManager"></param>
    internal ExcelConditionalFormattingIconSetBase(eExcelConditionalFormattingRuleType type,
                                                   ExcelAddress address,
                                                   int priority,
                                                   ExcelWorksheet worksheet,
                                                   XmlNode itemElementNode,
                                                   XmlNamespaceManager namespaceManager)
        : base(type, address, priority, worksheet, itemElementNode, namespaceManager == null ? worksheet.NameSpaceManager : namespaceManager)
    {
        if (itemElementNode != null && itemElementNode.HasChildNodes)
        {
            int pos = 1;

            foreach (XmlNode node in itemElementNode.SelectNodes("d:iconSet/d:cfvo", this.NameSpaceManager))
            {
                if (pos == 1)
                {
                    this.Icon1 = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, node, namespaceManager);
                }
                else if (pos == 2)
                {
                    this.Icon2 = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, node, namespaceManager);
                }
                else if (pos == 3)
                {
                    this.Icon3 = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, node, namespaceManager);
                }
                else
                {
                    break;
                }

                pos++;
            }
        }
        else
        {
            XmlNode? iconSetNode = this.CreateComplexNode(this.Node, ExcelConditionalFormattingConstants.Paths.IconSet);

            //Create the <iconSet> node inside the <cfRule> node
            double spann;

            if (type == eExcelConditionalFormattingRuleType.ThreeIconSet)
            {
                spann = 3;
            }
            else if (type == eExcelConditionalFormattingRuleType.FourIconSet)
            {
                spann = 4;
            }
            else
            {
                spann = 5;
            }

            XmlElement? iconNode1 = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
            _ = iconSetNode.AppendChild(iconNode1);

            this.Icon1 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent,
                                                                        0,
                                                                        "",
                                                                        eExcelConditionalFormattingRuleType.ThreeIconSet,
                                                                        address,
                                                                        priority,
                                                                        worksheet,
                                                                        iconNode1,
                                                                        namespaceManager);

            XmlElement? iconNode2 = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
            _ = iconSetNode.AppendChild(iconNode2);

            this.Icon2 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent,
                                                                        Math.Round(100D / spann, 0),
                                                                        "",
                                                                        eExcelConditionalFormattingRuleType.ThreeIconSet,
                                                                        address,
                                                                        priority,
                                                                        worksheet,
                                                                        iconNode2,
                                                                        namespaceManager);

            XmlElement? iconNode3 = iconSetNode.OwnerDocument.CreateElement(ExcelConditionalFormattingConstants.Paths.Cfvo, ExcelPackage.schemaMain);
            _ = iconSetNode.AppendChild(iconNode3);

            this.Icon3 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent,
                                                                        Math.Round(100D * (2D / spann), 0),
                                                                        "",
                                                                        eExcelConditionalFormattingRuleType.ThreeIconSet,
                                                                        address,
                                                                        priority,
                                                                        worksheet,
                                                                        iconNode3,
                                                                        namespaceManager);

            this.Type = type;
        }
    }

    /// <summary>
    /// 
    /// </summary>
    ///<param name="type"></param>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    internal ExcelConditionalFormattingIconSetBase(eExcelConditionalFormattingRuleType type,
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
    internal ExcelConditionalFormattingIconSetBase(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
        : this(type, address, priority, worksheet, null, null)
    {
    }

    #endregion Constructors

    /// <summary>
    /// Settings for icon 1 in the iconset
    /// </summary>
    public ExcelConditionalFormattingIconDataBarValue Icon1 { get; internal set; }

    /// <summary>
    /// Settings for icon 2 in the iconset
    /// </summary>
    public ExcelConditionalFormattingIconDataBarValue Icon2 { get; internal set; }

    /// <summary>
    /// Settings for icon 2 in the iconset
    /// </summary>
    public ExcelConditionalFormattingIconDataBarValue Icon3 { get; internal set; }

    private const string _reversePath = "d:iconSet/@reverse";

    /// <summary>
    /// Reverse the order of the icons
    /// </summary>
    public bool Reverse
    {
        get => this.GetXmlNodeBool(_reversePath, false);
        set => this.SetXmlNodeBool(_reversePath, value);
    }

    private const string _showValuePath = "d:iconSet/@showValue";

    /// <summary>
    /// If the cell values are visible
    /// </summary>
    public bool ShowValue
    {
        get => this.GetXmlNodeBool(_showValuePath, true);
        set => this.SetXmlNodeBool(_showValuePath, value);
    }

    private const string _iconSetPath = "d:iconSet/@iconSet";

    /// <summary>
    /// Type of iconset
    /// </summary>
    public T IconSet
    {
        get
        {
            string? v = this.GetXmlNodeString(_iconSetPath);
            v = v.Substring(1); //Skip first icon.

            return (T)Enum.Parse(typeof(T), v, true);
        }
        set => this.SetXmlNodeString(_iconSetPath, this.GetIconSetString(value));
    }

    private string GetIconSetString(T value)
    {
        if (this.Type == eExcelConditionalFormattingRuleType.FourIconSet)
        {
            switch (value.ToString())
            {
                case "Arrows":
                    return "4Arrows";

                case "ArrowsGray":
                    return "4ArrowsGray";

                case "Rating":
                    return "4Rating";

                case "RedToBlack":
                    return "4RedToBlack";

                case "TrafficLights":
                    return "4TrafficLights";

                default:
                    throw new ArgumentException("Invalid type");
            }
        }
        else if (this.Type == eExcelConditionalFormattingRuleType.FiveIconSet)
        {
            switch (value.ToString())
            {
                case "Arrows":
                    return "5Arrows";

                case "ArrowsGray":
                    return "5ArrowsGray";

                case "Quarters":
                    return "5Quarters";

                case "Rating":
                    return "5Rating";

                default:
                    throw new ArgumentException("Invalid type");
            }
        }
        else
        {
            switch (value.ToString())
            {
                case "Arrows":
                    return "3Arrows";

                case "ArrowsGray":
                    return "3ArrowsGray";

                case "Flags":
                    return "3Flags";

                case "Signs":
                    return "3Signs";

                case "Symbols":
                    return "3Symbols";

                case "Symbols2":
                    return "3Symbols2";

                case "TrafficLights1":
                    return "3TrafficLights1";

                case "TrafficLights2":
                    return "3TrafficLights2";

                default:
                    throw new ArgumentException("Invalid type");
            }
        }
    }
}