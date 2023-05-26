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
namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Differential formatting record used in conditional formatting
    /// </summary>
    public class ExcelDxfStyleFont : ExcelDxfStyleBase
    {
        internal ExcelDxfStyleFont(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
            : base(nameSpaceManager, topNode, styles, callback)
        {
            this.Font = new ExcelDxfFont(this._styles, callback);
            if (topNode != null)
            {
                this.Font.SetValuesFromXml(this._helper);
            }
        }
        /// <summary>
        /// Font formatting settings
        /// </summary>
        public ExcelDxfFont Font { get; internal set; }
        internal override string Id
        {
            get
            {
                return base.Id + this.Font.Id;
            }
        }
        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get
            {
                return base.HasValue || this.Font.HasValue;
            }
        }
        internal override DxfStyleBase Clone()
        {
            ExcelDxfStyleFont? s = new ExcelDxfStyleFont(this._helper.NameSpaceManager, null, this._styles, this._callback)
            {
                Font = (ExcelDxfFont)this.Font.Clone(),
                Fill = (ExcelDxfFill)this.Fill.Clone(),
                Border = (ExcelDxfBorderBase)this.Border.Clone(),
            };

            return s;
        }
        internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (this.Font.HasValue)
            {
                this.Font.CreateNodes(helper, "d:font");
            }

            if (this.Fill.HasValue)
            {
                this.Fill.CreateNodes(helper, "d:fill");
            }

            if (this.Border.HasValue)
            {
                this.Border.CreateNodes(helper, "d:border");
            }
        }
        internal override void SetStyle()
        {
            if (this._callback != null)
            {
                base.SetStyle();
                this.Font.SetStyle();
            }
        }
        /// <summary>
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {
            base.Clear();
            this.Font.Clear();
       }
    }
}
