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

using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Settings;

/// <summary>
/// Settings for html export for tables
/// </summary>
public class HtmlTableExportSettings : HtmlExportSettings
{
    /// <summary>
    /// Css export settings.
    /// </summary>
    public CssTableExportSettings Css { get; } = new CssTableExportSettings();

    /// <summary>
    /// Reset the settings to it's default values.
    /// </summary>
    public void ResetToDefault()
    {
        this.Minify = true;
        this.HiddenRows = eHiddenState.Exclude;
        this.Accessibility.TableSettings.ResetToDefault();
        this.IncludeCssClassNames = true;
        this.TableId = "";
        this.AdditionalTableClassNames.Clear();
        this.Culture = CultureInfo.CurrentCulture;
        this.Encoding = Encoding.UTF8;
        this.RenderDataAttributes = true;
        this.Css.ResetToDefault();
        this.Pictures.ResetToDefault();
    }

    /// <summary>
    /// Copy the values from another settings object.
    /// </summary>
    /// <param name="copy">The object to copy.</param>
    public void Copy(HtmlTableExportSettings copy)
    {
        this.Minify = copy.Minify;
        this.HiddenRows = copy.HiddenRows;
        this.Accessibility.TableSettings.Copy(copy.Accessibility.TableSettings);
        this.IncludeCssClassNames = copy.IncludeCssClassNames;
        this.TableId = copy.TableId;
        this.AdditionalTableClassNames = copy.AdditionalTableClassNames;
        this.Culture = copy.Culture;
        this.Encoding = copy.Encoding;
        this.RenderDataAttributes = copy.RenderDataAttributes;
        this.Css.Copy(copy.Css);
        this.Pictures.Copy(copy.Pictures);
    }

    /// <summary>
    /// Configure the settings.
    /// </summary>
    /// <param name="settings"></param>
    public void Configure(Action<HtmlTableExportSettings> settings) => settings.Invoke(this);
}