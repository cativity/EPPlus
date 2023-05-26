/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
 Date               Author                       Change
 *************************************************************************************************
 12/28/2020         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Base class for differential formatting styles
    /// </summary>
    public abstract class ExcelDxfStyleBase : DxfStyleBase 
    {
        internal XmlHelperInstance _helper;            
        //internal protected string _dxfIdPath;

        internal ExcelDxfStyleBase(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback) : base(styles, callback)
        {
            //_dxfIdPath = dxfIdPath;
            this.Border = new ExcelDxfBorderBase(this._styles, callback);
            this.Fill = new ExcelDxfFill(this._styles, callback);

            if (topNode != null)
            {
                this._helper = new XmlHelperInstance(nameSpaceManager, topNode);
                this.Border.SetValuesFromXml(this._helper);
                this.Fill.SetValuesFromXml(this._helper);
            }
            else
            {
                this._helper = new XmlHelperInstance(nameSpaceManager);
            }

            this._helper.SchemaNodeOrder = new string[] { "font", "numFmt", "fill", "border" };
        }
        internal virtual int DxfId { get; set; } = int.MinValue;
        /// <summary>
        /// Fill formatting settings
        /// </summary>
        public ExcelDxfFill Fill { get; set; }
        /// <summary>
        /// Border formatting settings
        /// </summary>
        public ExcelDxfBorderBase Border { get; set; }
        /// <summary>
        /// Id
        /// </summary>
        internal override string Id
        {
            get
            {
                return this.Border.Id + this.Fill.Id +
                       (this.AllowChange ? "" : this.DxfId.ToString());
            }
        }
        
        /// <summary>
        /// Creates the node
        /// </summary>
        /// <param name="helper">The helper</param>
        /// <param name="path">The XPath</param>
        internal override void CreateNodes(XmlHelper helper, string path)
        {
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
                this.Border.SetStyle();
                this.Fill.SetStyle();
            }
        }

        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get 
            {
                return this.Fill.HasValue || this.Border.HasValue; 
            }
        }
        /// <summary>
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {
            this.Fill.Clear();
            this.Border.Clear();
        }
        internal ExcelDxfStyle ToDxfStyle()
        {
            if (this is ExcelDxfStyle s)
            {
                return s;
            }
            else
            {
                ExcelDxfStyle? ns = new ExcelDxfStyle(this._styles.NameSpaceManager, null, this._styles, null)
                {
                    Border = this.Border,
                    Fill = this.Fill,
                    DxfId = this.DxfId,
                    Font = new ExcelDxfFont(this._styles, this._callback),
                    NumberFormat = new ExcelDxfNumberFormat(this._styles, this._callback),
                    _helper = this._helper
                };
                ns.Font.SetValuesFromXml(this._helper);
                return ns;
            }
        }
        internal ExcelDxfSlicerStyle ToDxfSlicerStyle()
        {
            if (this is ExcelDxfSlicerStyle s)
            {
                return s;
            }
            else
            {
                ExcelDxfSlicerStyle? ns = new ExcelDxfSlicerStyle(this._styles.NameSpaceManager, null, this._styles, null)
                {
                    Border = this.Border,
                    Fill = this.Fill,
                    DxfId = this.DxfId,
                    Font = new ExcelDxfFont(this._styles, this._callback),
                    _helper = this._helper
                };
                ns.Font.SetValuesFromXml(this._helper);
                return ns;
            }
        }
        internal ExcelDxfTableStyle ToDxfTableStyle()
        {
            if(this is ExcelDxfTableStyle s)
            {
                return s;
            }
            else
            {
                ExcelDxfTableStyle? ns = new ExcelDxfTableStyle(this._styles.NameSpaceManager, null, this._styles)
                {
                    Border = this.Border,
                    Fill = this.Fill,
                    DxfId = this.DxfId,
                    Font = new ExcelDxfFont(this._styles, this._callback),
                    _helper = this._helper
                };
                ns.Font.SetValuesFromXml(this._helper);
                return ns;
            }
        }
        internal ExcelDxfStyleLimitedFont ToDxfLimitedStyle()
        {
            if (this is ExcelDxfStyleLimitedFont s)
            {
                return s;
            }
            else
            {
                ExcelDxfStyleLimitedFont? ns = new ExcelDxfStyleLimitedFont(this._styles.NameSpaceManager, null, this._styles, this._callback)
                {
                    Border = this.Border,
                    Fill = this.Fill,
                    DxfId = this.DxfId,
                    Font = new ExcelDxfFontBase(this._styles, this._callback),
                    _helper = this._helper
                };
                ns.Font.SetValuesFromXml(this._helper);
                return ns;
            }
        }
        internal ExcelDxfStyleConditionalFormatting ToDxfConditionalFormattingStyle()
        {
            if (this is ExcelDxfStyleConditionalFormatting s)
            {
                return s;
            }
            else
            {
                ExcelDxfStyleConditionalFormatting? ns = new ExcelDxfStyleConditionalFormatting(this._styles.NameSpaceManager, null, this._styles, this._callback)
                {
                    Border = this.Border,
                    Fill = this.Fill,
                    NumberFormat = new ExcelDxfNumberFormat(this._styles, this._callback),
                    DxfId = this.DxfId,
                    Font = new ExcelDxfFontBase(this._styles, this._callback),
                    _helper = this._helper
                };
                ns.NumberFormat.SetValuesFromXml(this._helper);
                ns.Font.SetValuesFromXml(this._helper);
                return ns;
            }
        }
    }
}
