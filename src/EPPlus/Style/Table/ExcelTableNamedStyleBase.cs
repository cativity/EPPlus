/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/

using OfficeOpenXml.Core;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Style.Table;

/// <summary>
/// A base class for custom named table styles
/// </summary>
public abstract class ExcelTableNamedStyleBase : XmlHelper
{
    internal ExcelStyles _styles;
    internal Dictionary<eTableStyleElement, ExcelTableStyleElement> _dic = new Dictionary<eTableStyleElement, ExcelTableStyleElement>();

    internal ExcelTableNamedStyleBase(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles)
        : base(nameSpaceManager, topNode)
    {
        this._styles = styles;
        this.As = new ExcelTableNamedStyleAsType(this);

        foreach (XmlNode node in topNode.ChildNodes)
        {
            if (node is XmlElement e)
            {
                eTableStyleElement type = e.GetAttribute("type").ToEnum(eTableStyleElement.WholeTable);

                if (IsBanded(type))
                {
                    this._dic.Add(type, new ExcelBandedTableStyleElement(nameSpaceManager, node, styles, type));
                }
                else
                {
                    this._dic.Add(type, new ExcelTableStyleElement(nameSpaceManager, node, styles, type));
                }
            }
        }
    }

    internal static bool IsBanded(eTableStyleElement type) =>
        type == eTableStyleElement.FirstColumnStripe
        || type == eTableStyleElement.FirstRowStripe
        || type == eTableStyleElement.SecondColumnStripe
        || type == eTableStyleElement.SecondRowStripe;

    internal ExcelTableStyleElement GetTableStyleElement(eTableStyleElement element)
    {
        if (this._dic.ContainsKey(element))
        {
            return this._dic[element];
        }

        ExcelTableStyleElement item;

        if (IsBanded(element))
        {
            item = new ExcelBandedTableStyleElement(this.NameSpaceManager, this.TopNode, this._styles, element);
        }
        else
        {
            item = new ExcelTableStyleElement(this.NameSpaceManager, this.TopNode, this._styles, element);
        }

        this._dic.Add(element, item);

        return item;
    }

    /// <summary>
    /// If a table style is applied for a table/pivot table or both
    /// </summary>
    public abstract eTableNamedStyleAppliesTo AppliesTo { get; }

    /// <summary>
    /// The name of the table named style
    /// </summary>
    public string Name
    {
        get => this.GetXmlNodeString("@name");
        set
        {
            if (this._styles.TableStyles.ExistsKey(value) || this._styles.SlicerStyles.ExistsKey(value))
            {
                throw new InvalidOperationException("Name already is already used by a table or slicer style");
            }

            this.SetXmlNodeString("@name", value);
        }
    }

    /// <summary>
    /// Applies to the entire content of a table or pivot table
    /// </summary>
    public ExcelTableStyleElement WholeTable => this.GetTableStyleElement(eTableStyleElement.WholeTable);

    /// <summary>
    /// Applies to the first column stripe of a table or pivot table
    /// </summary>
    public ExcelBandedTableStyleElement FirstColumnStripe => (ExcelBandedTableStyleElement)this.GetTableStyleElement(eTableStyleElement.FirstColumnStripe);

    /// <summary>
    /// Applies to the second column stripe of a table or pivot table
    /// </summary>
    public ExcelBandedTableStyleElement SecondColumnStripe => (ExcelBandedTableStyleElement)this.GetTableStyleElement(eTableStyleElement.SecondColumnStripe);

    /// <summary>
    /// Applies to the first row stripe of a table or pivot table
    /// </summary>
    public ExcelBandedTableStyleElement FirstRowStripe => (ExcelBandedTableStyleElement)this.GetTableStyleElement(eTableStyleElement.FirstRowStripe);

    /// <summary>
    /// Applies to the second row stripe of a table or pivot table
    /// </summary>
    public ExcelBandedTableStyleElement SecondRowStripe => (ExcelBandedTableStyleElement)this.GetTableStyleElement(eTableStyleElement.SecondRowStripe);

    /// <summary>
    /// Applies to the last column of a table or pivot table
    /// </summary>
    public ExcelTableStyleElement LastColumn => this.GetTableStyleElement(eTableStyleElement.LastColumn);

    /// <summary>
    /// Applies to the first column of a table or pivot table
    /// </summary>
    public ExcelTableStyleElement FirstColumn => this.GetTableStyleElement(eTableStyleElement.FirstColumn);

    /// <summary>
    /// Applies to the header row of a table or pivot table
    /// </summary>
    public ExcelTableStyleElement HeaderRow => this.GetTableStyleElement(eTableStyleElement.HeaderRow);

    /// <summary>
    /// Applies to the total row of a table or pivot table
    /// </summary>
    public ExcelTableStyleElement TotalRow => this.GetTableStyleElement(eTableStyleElement.TotalRow);

    /// <summary>
    /// Applies to the first header cell of a table or pivot table
    /// </summary>
    public ExcelTableStyleElement FirstHeaderCell => this.GetTableStyleElement(eTableStyleElement.FirstHeaderCell);

    /// <summary>
    /// Provides access to type conversion for all table named styles.
    /// </summary>
    public ExcelTableNamedStyleAsType As { get; }

    internal void SetFromTemplate(ExcelTableNamedStyleBase templateStyle)
    {
        foreach (ExcelTableStyleElement? s in templateStyle._dic.Values)
        {
            ExcelTableStyleElement? element = this.GetTableStyleElement(s.Type);
            element.Style = (ExcelDxfStyleLimitedFont)s.Style.Clone();
        }
    }

    internal void SetFromTemplate(TableStyles templateStyle) => this.LoadTableTemplate("TableStyles", templateStyle.ToString());

    internal void SetFromTemplate(PivotTableStyles templateStyle) => this.LoadTableTemplate("PivotTableStyles", templateStyle.ToString());

    private void LoadTableTemplate(string folder, string styleName)
    {
        ZipInputStream? zipStream = ZipHelper.OpenZipResource();

        while (zipStream.GetNextEntry() is { } entry)
        {
            if (entry.IsDirectory || !entry.FileName.EndsWith(".xml") || entry.UncompressedSize <= 0 || !entry.FileName.StartsWith(folder))
            {
                continue;
            }

            string? name = new FileInfo(entry.FileName).Name;
            name = name.Substring(0, name.Length - 4);

            if (name.Equals(styleName, StringComparison.OrdinalIgnoreCase))
            {
                string? xmlContent = ZipHelper.UncompressEntry(zipStream, entry);
                XmlDocument? xml = new XmlDocument();
                xml.LoadXml(xmlContent);

                foreach (XmlElement elem in xml.DocumentElement.ChildNodes)
                {
                    eTableStyleElement type = elem.GetAttribute("name").ToEnum(eTableStyleElement.WholeTable);
                    ExcelDxfStyleLimitedFont? dxf = new ExcelDxfStyleLimitedFont(this.NameSpaceManager, elem.FirstChild, this._styles, null);

                    ExcelTableStyleElement? te = this.GetTableStyleElement(type);
                    te.Style = dxf;
                }
            }
        }

        if (string.IsNullOrEmpty(this.Name))
        {
            this.SetXmlNodeString("@name", styleName);
        }
    }
}