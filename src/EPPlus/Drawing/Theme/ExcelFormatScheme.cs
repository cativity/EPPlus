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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// The background fill styles, effect styles, fill styles, and line styles which define the style matrix for a theme
    /// </summary>
    public class ExcelFormatScheme : XmlHelper
    {

        private readonly ExcelThemeBase _theme;
        internal ExcelFormatScheme(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelThemeBase theme) : base(nameSpaceManager, topNode)
        {
            this._theme = theme;
        }
        /// <summary>
        /// The name of the format scheme
        /// </summary>
        public string Name
        {
            get
            {
                return this.GetXmlNodeString("@name");
            }
            set
            {
                this.SetXmlNodeString("@name", value);
            }
        }
        const string fillStylePath = "a:fillStyleLst";
        ExcelThemeFillStyles _fillStyle = null;
        /// <summary>
        ///  Defines the fill styles for the theme
        /// </summary>
        public ExcelThemeFillStyles FillStyle
        {
            get
            {
                if (this._fillStyle == null)
                {
                    this._fillStyle = new ExcelThemeFillStyles(this.NameSpaceManager, this.TopNode.SelectSingleNode(fillStylePath, this.NameSpaceManager), this._theme);
                }
                return this._fillStyle;
            }
        }

        const string lineStylePath = "a:lnStyleLst";
        ExcelThemeLineStyles _lineStyle = null;
        /// <summary>
        ///  Defines the line styles for the theme
        /// </summary>
        public ExcelThemeLineStyles BorderStyle
        {
            get
            {
                if (this._lineStyle == null)
                {
                    this._lineStyle = new ExcelThemeLineStyles(this.NameSpaceManager, this.TopNode.SelectSingleNode(lineStylePath, this.NameSpaceManager));
                }
                return this._lineStyle;
            }
        }
        const string effectStylePath = "a:effectStyleLst";
        ExcelThemeEffectStyles _effectStyle = null;
        /// <summary>
        ///  Defines the effect styles for the theme
        /// </summary>
        public ExcelThemeEffectStyles EffectStyle
        {
            get
            {
                if (this._effectStyle == null)
                {
                    this._effectStyle = new ExcelThemeEffectStyles(this.NameSpaceManager, this.TopNode.SelectSingleNode(effectStylePath, this.NameSpaceManager), this._theme);
                }
                return this._effectStyle;
            }
        }
        const string backgroundFillStylePath = "a:bgFillStyleLst";
        ExcelThemeFillStyles _backgroundFillStyle = null;
        /// <summary>
        /// Define background fill styles for the theme
        /// </summary>
        public ExcelThemeFillStyles BackgroundFillStyle
        {
            get
            {
                if (this._backgroundFillStyle == null)
                {
                    this._backgroundFillStyle = new ExcelThemeFillStyles(this.NameSpaceManager, this.TopNode.SelectSingleNode(backgroundFillStylePath, this.NameSpaceManager), this._theme);
                }
                return this._backgroundFillStyle;
            }
        }
    }
}
