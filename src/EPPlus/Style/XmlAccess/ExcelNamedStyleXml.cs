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
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for named styles
    /// </summary>
    public sealed class ExcelNamedStyleXml : StyleXmlHelper
    {
        ExcelStyles _styles;
        internal ExcelNamedStyleXml(XmlNamespaceManager nameSpaceManager, ExcelStyles styles)
            : base(nameSpaceManager)
        {
            this._styles = styles;
            this.BuildInId = int.MinValue;
        }
        internal ExcelNamedStyleXml(XmlNamespaceManager NameSpaceManager, XmlNode topNode, ExcelStyles styles) :
            base(NameSpaceManager, topNode)
        {
            this.StyleXfId = this.GetXmlNodeInt(idPath);
            this.Name = this.GetXmlNodeString(namePath);
            this.BuildInId = this.GetXmlNodeInt(buildInIdPath);
            this.CustomBuildin = this.GetXmlNodeBool(customBuiltinPath);
            this.Uid= this.GetXmlNodeString(uidPath);
            this._styles = styles;
            this._style = new ExcelStyle(styles, styles.NamedStylePropertyChange, -1, this.Name, this._styleXfId);
        }
        internal override string Id
        {
            get
            {
                return this.Name;
            }
        }
        int _styleXfId=0;
        const string idPath = "@xfId";
        /// <summary>
        /// Named style index
        /// </summary>
        public int StyleXfId
        {
            get
            {
                return this._styleXfId;
            }
            set
            {
                this._styleXfId = value;
            }
        }
        int _xfId = int.MinValue;
        /// <summary>
        /// Style index
        /// </summary>
        internal int XfId
        {
            get
            {
                return this._xfId;
            }
            set
            {
                this._xfId = value;
            }
        }
        const string buildInIdPath = "@builtinId";
        /// <summary>
        /// The build in Id for the named style
        /// </summary>
        public int BuildInId { get; set; }
        const string customBuiltinPath = "@customBuiltin";
        /// <summary>
        /// Indicates if this built-in cell style has been customized
        /// </summary>
        public bool CustomBuildin { get; set; }
        const string namePath = "@name";
        string _name;
        /// <summary>
        /// Name of the style
        /// </summary>
        public string Name
        {
            get
            {
                return this._name;
            }
            internal set
            {
                this._name = value;
            }
        }
        ExcelStyle _style = null;
        /// <summary>
        /// The style object
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return this._style;
            }
            internal set
            {
                this._style = value;
            }
        }
        const string uidPath="@xr:uid";
        internal string Uid { get; set; }
        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            this.TopNode = topNode;
            this.SetXmlNodeString(namePath, this._name);
            this.SetXmlNodeString(idPath, this._styles.CellStyleXfs[this.StyleXfId].newID.ToString());
            if (this.BuildInId>=0)
            {
                this.SetXmlNodeString(buildInIdPath, this.BuildInId.ToString());
            }

            if(this.CustomBuildin)
            {
                this.SetXmlNodeBool(customBuiltinPath, true);
            }

            if (!string.IsNullOrEmpty(this.Uid))
            {
                this.SetXmlNodeString(uidPath, this.Uid);
            }

            return this.TopNode;            
        }
    }
}
