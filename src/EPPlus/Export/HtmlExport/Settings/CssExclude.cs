﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/11/2021         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/

namespace OfficeOpenXml.Export.HtmlExport;

/// <summary>
/// Css settings to exclude individual styles.
/// </summary>
public class CssExclude
{
    internal CssExclude()
    {
    }

    /// <summary>
    /// Exclude Font styles.
    /// </summary>
    public eFontExclude Font { get; set; }

    /// <summary>
    /// Exclude Border styles
    /// </summary>
    public eBorderExclude Border { get; set; }

    /// <summary>
    /// Exclude Fill styles
    /// </summary>
    public bool Fill { get; set; }

    /// <summary>
    /// Exclude vertical alignment.
    /// </summary>
    public bool VerticalAlignment { get; set; }

    /// <summary>
    /// Exclude horizontal alignment.
    /// </summary>
    public bool HorizontalAlignment { get; set; }

    /// <summary>
    /// Exclude Wrap Text
    /// </summary>
    public bool WrapText { get; set; }

    /// <summary>
    /// Exclude Text Rotation
    /// </summary>
    public bool TextRotation { get; set; }

    /// <summary>
    /// Exclude Indent.
    /// </summary>
    public bool Indent { get; set; }

    /// <summary>
    /// Reset the settings to it's default values.
    /// </summary>
    public void ResetToDefault()
    {
        this.Font = 0;
        this.Border = 0;
        this.Fill = false;
        this.VerticalAlignment = false;
        this.HorizontalAlignment = false;
        this.WrapText = false;
        this.TextRotation = false;
        this.Indent = false;
    }

    /// <summary>
    /// Copy the values from another settings object.
    /// </summary>
    /// <param name="copy">The object to copy.</param>
    public void Copy(CssExclude copy)
    {
        this.Font = copy.Font;
        this.Border = copy.Border;
        this.Fill = copy.Fill;
        this.VerticalAlignment = copy.VerticalAlignment;
        this.HorizontalAlignment = copy.HorizontalAlignment;
        this.WrapText = copy.WrapText;
        this.TextRotation = copy.TextRotation;
        this.Indent = copy.Indent;
    }
}