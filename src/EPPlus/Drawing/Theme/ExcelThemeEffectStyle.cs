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
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// An effect style for a theme
    /// </summary>
    public class ExcelThemeEffectStyle : XmlHelper
    {
        string _path;
        string[] _schemaNodeOrder;
        private readonly ExcelThemeBase _theme;
        internal ExcelThemeEffectStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder, ExcelThemeBase theme) : base(nameSpaceManager, topNode)
        {
            if (!string.IsNullOrEmpty(path))
            {
                path += "/";
            }

            this._path = path;
            this._schemaNodeOrder = schemaNodeOrder;
            this._theme = theme;
        }
        ExcelDrawingEffectStyle _effects = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if(this._effects==null)
                {
                    this._effects = new ExcelDrawingEffectStyle(this._theme, this.NameSpaceManager, this.TopNode, this._path + "a:effectLst", this._schemaNodeOrder);
                }
                return this._effects;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D settings
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (this._threeD == null)
                {
                    this._threeD = new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, this._path, this._schemaNodeOrder);
                }
                return this._threeD;
            }
        }
    }
}
