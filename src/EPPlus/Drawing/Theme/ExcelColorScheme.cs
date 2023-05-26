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
using System.Xml;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Style;
namespace OfficeOpenXml.Drawing.Theme
{

    /// <summary>
    /// The color Scheme for a theme
    /// </summary>
    public class ExcelColorScheme : XmlHelper
    {
        internal ExcelColorScheme(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            this.SchemaNodeOrder = new string[] { "dk1","lt1", "dk2", "lt3", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
        }
        const string Dk1Path = "a:dk1";
        ExcelDrawingThemeColorManager _dk1 =null;
        /// <summary>
        /// Dark 1 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Dark1
        {
            get { return this._dk1 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, Dk1Path, this.SchemaNodeOrder); }
        }

        internal ExcelDrawingThemeColorManager GetColorByEnum(eThemeSchemeColor color)
        {
            switch(color)
            {
                case eThemeSchemeColor.Accent1:
                    return this.Accent1;
                case eThemeSchemeColor.Accent2:
                    return this.Accent2;
                case eThemeSchemeColor.Accent3:
                    return this.Accent3;
                case eThemeSchemeColor.Accent4:
                    return this.Accent4;
                case eThemeSchemeColor.Accent5:
                    return this.Accent5;
                case eThemeSchemeColor.Accent6:
                    return this.Accent6;
                case eThemeSchemeColor.Background1:
                    return this.Light1;
                case eThemeSchemeColor.Background2:
                    return this.Light2;
                case eThemeSchemeColor.Text1:
                    return this.Dark1;
                case eThemeSchemeColor.Text2:
                    return this.Dark2;
                case eThemeSchemeColor.Hyperlink:
                    return this.Hyperlink;
                case eThemeSchemeColor.FollowedHyperlink:
                    return this.FollowedHyperlink;                
            }
            throw(new ArgumentOutOfRangeException($"Type {color} is unhandled."));
        }

        const string Dk2Path = "a:dk2";
        ExcelDrawingThemeColorManager _dk2 = null;
        /// <summary>
        /// Dark 2 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Dark2
        {
            get { return this._dk2 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, Dk2Path, this.SchemaNodeOrder); }
        }
        const string lt1Path = "a:lt1";
        ExcelDrawingThemeColorManager _lt1 = null;
        /// <summary>
        /// Light 1 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Light1
        {
            get { return this._lt1 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, lt1Path, this.SchemaNodeOrder); }
        }
        const string lt2Path = "a:lt2";
        ExcelDrawingThemeColorManager _lt2 = null;
        /// <summary>
        /// Light 2 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Light2
        {
            get { return this._lt2 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, lt2Path, this.SchemaNodeOrder); }
        }
        const string Accent1Path = "a:accent1";
        ExcelDrawingThemeColorManager _accent1 = null;
        /// <summary>
        /// Accent 1 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent1
        {
            get
            {
                return this._accent1 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, Accent1Path, this.SchemaNodeOrder);
            }
        }
        const string Accent2Path = "a:accent2";
        ExcelDrawingThemeColorManager _accent2 = null;
        /// <summary>
        /// Accent 2 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent2
        {
            get
            {
                return this._accent2 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, Accent2Path, this.SchemaNodeOrder);
            }
        }
        const string Accent3Path = "a:accent3";
        ExcelDrawingThemeColorManager _accent3 = null;
        /// <summary>
        /// Accent 3 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent3
        {
            get
            {
                return this._accent3 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, Accent3Path, this.SchemaNodeOrder);
            }
        }
        const string Accent4Path = "a:accent4";
        ExcelDrawingThemeColorManager _accent4 = null;
        /// <summary>
        /// Accent 4 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent4
        {
            get
            {
                return this._accent4 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, Accent4Path, this.SchemaNodeOrder);
            }
        }
        const string Accent5Path = "a:accent5";
        ExcelDrawingThemeColorManager _accent5 = null;
        /// <summary>
        /// Accent 5 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent5
        {
            get
            {
                return this._accent5 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, Accent5Path, this.SchemaNodeOrder);
            }
        }
        const string Accent6Path = "a:accent6";
        ExcelDrawingThemeColorManager _accent6 = null;
        /// <summary>
        /// Accent 6 theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Accent6
        {
            get
            {
                return this._accent6 ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, Accent6Path, this.SchemaNodeOrder);
            }
        }
        const string HlinkPath = "a:hlink";
        ExcelDrawingThemeColorManager _hlink = null;
        /// <summary>
        /// Hyperlink theme color
        /// </summary>
        public ExcelDrawingThemeColorManager Hyperlink
        {
            get
            {
                return this._hlink ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, HlinkPath, this.SchemaNodeOrder);
            }
        }

        const string FolHlinkPath = "a:folHlink";
        ExcelDrawingThemeColorManager _folHlink = null;
        /// <summary>
        /// Followed hyperlink theme color
        /// </summary>
        public ExcelDrawingThemeColorManager FollowedHyperlink
        {
            get
            {
                return this._folHlink ??= new ExcelDrawingThemeColorManager(this.NameSpaceManager, this.TopNode, FolHlinkPath, this.SchemaNodeOrder);
            }
        }
    }
}
