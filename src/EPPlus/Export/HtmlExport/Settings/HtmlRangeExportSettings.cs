/*************************************************************************************************
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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport;

/// <summary>
/// Settings for html export for ranges
/// </summary>
public class HtmlRangeExportSettings : HtmlExportSettings
{
    int _headerRows = 1;

    /// <summary>
    /// Number of header rows before the actual data. Default is 1.
    /// </summary>
    public int HeaderRows
    {
        get => this._headerRows;
        set
        {
            if (value < 0 || value > ExcelPackage.MaxRows)
            {
                throw new InvalidOperationException("Can't be negative or exceed number of allowed rows in a worksheet.");
            }

            this._headerRows = value;
        }
    }

    /// <summary>
    /// If <see cref="HeaderRows"/> is 0, this collection contains the headers. 
    /// If this collection is empty the table will have no headers.
    /// </summary>
    public List<string> Headers { get; } = new List<string>();

    /// <summary>
    /// Options to exclude css elements
    /// </summary>
    public CssRangeExportSettings Css { get; } = new CssRangeExportSettings();

    /// <summary>
    /// Reset the setting to it's default values.
    /// </summary>
    public void ResetToDefault()
    {
        this.Minify = true;
        this.HiddenRows = eHiddenState.Exclude;
        this.HeaderRows = 1;
        this.Headers.Clear();
        this.Accessibility.TableSettings.ResetToDefault();
        this.AdditionalTableClassNames.Clear();
        this.Culture = CultureInfo.CurrentCulture;
        this.Encoding = Encoding.UTF8;
        this.Css.ResetToDefault();
        this.Pictures.ResetToDefault();
    }

    /// <summary>
    /// Copy the values from another settings object.
    /// </summary>
    /// <param name="copy">The object to copy.</param>
    public void Copy(HtmlRangeExportSettings copy)
    {
        this.Minify = copy.Minify;
        this.HiddenRows = copy.HiddenRows;
        this.HeaderRows = copy.HeaderRows;
        this.Headers.Clear();
        this.Headers.AddRange(copy.Headers);

        this.Accessibility.TableSettings.Copy(copy.Accessibility.TableSettings);

        this.AdditionalTableClassNames.Clear();
        this.AdditionalTableClassNames.AddRange(copy.AdditionalTableClassNames);

        this.Culture = copy.Culture;
        this.Encoding = copy.Encoding;
        this.Css.Copy(copy.Css);
        this.Pictures.Copy(copy.Pictures);
    }
}