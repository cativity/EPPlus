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
  02/26/2021         EPPlus Software AB       Modified to work with dxf styling for tables
 *************************************************************************************************/
using System;
using System.Xml;
namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// Differential formatting record used in conditional formatting
/// </summary>
public class ExcelDxfStyleConditionalFormatting : ExcelDxfStyleLimitedFont
{
    internal ExcelDxfStyleConditionalFormatting(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
        : base(nameSpaceManager, topNode, styles, callback)
    {
        this.NumberFormat = new ExcelDxfNumberFormat(this._styles, callback);
        if (topNode != null)
        {
            this.NumberFormat.SetValuesFromXml(this._helper);
        }
    }
    /// <summary>
    /// Number format settings
    /// </summary>
    public ExcelDxfNumberFormat NumberFormat { get; internal set; }
    internal override string Id
    {
        get
        {
            return base.Id + this.NumberFormat.Id;
        }
    }
    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue
    {
        get
        {
            return base.HasValue || this.NumberFormat.HasValue;
        }
    }
    internal override DxfStyleBase Clone()
    {
        ExcelDxfStyleConditionalFormatting? s = new ExcelDxfStyleConditionalFormatting(this._helper.NameSpaceManager, null, this._styles, this._callback)
        {
            Font = (ExcelDxfFontBase)this.Font.Clone(),
            Fill = (ExcelDxfFill)this.Fill.Clone(),
            Border = (ExcelDxfBorderBase)this.Border.Clone(),
            NumberFormat = (ExcelDxfNumberFormat)this.NumberFormat.Clone(),
        };

        return s;
    }
    internal override void CreateNodes(XmlHelper helper, string path)
    {
        if (this.Font.HasValue)
        {
            this.Font.CreateNodes(helper, "d:font");
        }

        if (this.NumberFormat.HasValue)
        {
            this.NumberFormat.CreateNodes(helper, "d:numFmt");
        }

        base.CreateNodes(helper, path);
    }
    internal override void SetStyle()
    {
        if (this._callback != null)
        {
            base.SetStyle();
            this.NumberFormat.SetStyle();
        }
    }
    /// <summary>
    /// Clears all properties
    /// </summary>
    public override void Clear()
    {            
        base.Clear();
        this.NumberFormat.Clear();
    }
}