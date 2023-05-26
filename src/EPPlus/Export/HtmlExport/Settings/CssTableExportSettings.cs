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

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Settings for css export for tables
    /// </summary>
    public class CssTableExportSettings : CssExportSettings
    {
        internal CssTableExportSettings()
        {
            this.ResetToDefault();
        }
        /// <summary>
        /// Include Css for the current table style
        /// </summary>
        public bool IncludeTableStyles { get; set; } = true;
        /// <summary>
        /// Include Css for cell styling.
        /// </summary>
        public bool IncludeCellStyles { get; set; } = true;
        /// <summary>
        /// Exclude flags for styles
        /// </summary>
        public CssExcludeStyle Exclude
        {
            get;
        } = new CssExcludeStyle();

        /// <summary>
        /// Reset the settings to it's default values.
        /// </summary>
        public void ResetToDefault()
        {
            this.IncludeTableStyles = true;
            this.IncludeCellStyles = true;

            this.Exclude.TableStyle.ResetToDefault();
            this.Exclude.CellStyle.ResetToDefault();
            this.ResetToDefaultInternal();
        }
        /// <summary>
        /// Copy the values from another settings object.
        /// </summary>
        /// <param name="copy">The object to copy.</param>
        public void Copy(CssTableExportSettings copy)
        {
            this.IncludeTableStyles = copy.IncludeTableStyles;
            this.IncludeCellStyles = copy.IncludeTableStyles;

            this.Exclude.TableStyle.Copy(copy.Exclude.TableStyle);
            this.Exclude.CellStyle.Copy(copy.Exclude.CellStyle);

            this.CopyInternal(copy);
        }
    }
}
