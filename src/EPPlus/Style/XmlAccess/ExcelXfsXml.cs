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
using System.Globalization;
using System.Text;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class xfs records. This is the top level style object.
    /// </summary>
    public sealed class ExcelXfs : StyleXmlHelper
    {
        private readonly ExcelStyles _styles;
        internal ExcelXfs(XmlNamespaceManager nameSpaceManager, ExcelStyles styles) : base(nameSpaceManager)
        {
            this._styles = styles;
            this.isBuildIn = false;
        }
        internal ExcelXfs(XmlNamespaceManager nsm, XmlNode topNode, ExcelStyles styles) :
            base(nsm, topNode)
        {
            this._styles = styles;
            this.XfId = this.GetXmlNodeInt("@xfId");
            if (this.XfId == 0)
            {
                this.isBuildIn = true; //Normal taggen
            }

            this._numFmtId = this.GetXmlNodeInt("@numFmtId");
            this.FontId = this.GetXmlNodeInt("@fontId");
            this.FillId = this.GetXmlNodeInt("@fillId");
            this.BorderId = this.GetXmlNodeInt("@borderId");
            this._readingOrder = GetReadingOrder(this.GetXmlNodeString(readingOrderPath));
            this._indent = this.GetXmlNodeInt(indentPath);
            this.ShrinkToFit = this.GetXmlNodeString(shrinkToFitPath) == "1" ? true : false;
            this.VerticalAlignment = GetVerticalAlign(this.GetXmlNodeString(verticalAlignPath));
            this.HorizontalAlignment = GetHorizontalAlign(this.GetXmlNodeString(horizontalAlignPath));
            this.WrapText = this.GetXmlNodeBool(wrapTextPath);
            this._textRotation = this.GetXmlNodeInt(this.textRotationPath);
            this.Hidden = this.GetXmlNodeBool(hiddenPath);
            this.Locked = this.GetXmlNodeBool(lockedPath,true);
            this.QuotePrefix = this.GetXmlNodeBool(quotePrefixPath);
            this.JustifyLastLine = this.GetXmlNodeBool(justifyLastLine);
            this.ApplyAlignment = this.GetXmlNodeBoolNullable("@applyAlignment");
            this.ApplyBorder = this.GetXmlNodeBoolNullable("@applyBorder");
            this.ApplyFill = this.GetXmlNodeBoolNullable("@applyFill");
            this.ApplyFont = this.GetXmlNodeBoolNullable("@applyFont");
            this.ApplyNumberFormat = this.GetXmlNodeBoolNullable("@applyNumberFormat");
            this.ApplyProtection = this.GetXmlNodeBoolNullable("@applyProtection");
        }

        private static ExcelReadingOrder GetReadingOrder(string value)
        {
            switch(value)
            {
                case "1":
                    return ExcelReadingOrder.LeftToRight;
                case "2":
                    return ExcelReadingOrder.RightToLeft;
                default:
                    return ExcelReadingOrder.ContextDependent;
            }
        }

        private static ExcelHorizontalAlignment GetHorizontalAlign(string align)
        {
            if (align == "")
            {
                return ExcelHorizontalAlignment.General;
            }

            align = align.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + align.Substring(1, align.Length - 1);
            try
            {
                return (ExcelHorizontalAlignment)Enum.Parse(typeof(ExcelHorizontalAlignment), align);
            }
            catch
            {
                return ExcelHorizontalAlignment.General;
            }
        }

        private static ExcelVerticalAlignment GetVerticalAlign(string align)
        {
            if (align == "")
            {
                return ExcelVerticalAlignment.Bottom;
            }

            align = align.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + align.Substring(1, align.Length - 1);
            try
            {
                return (ExcelVerticalAlignment)Enum.Parse(typeof(ExcelVerticalAlignment), align);
            }
            catch
            {
                return ExcelVerticalAlignment.Bottom;
            }
        }
        /// <summary>
        /// Style index
        /// </summary>
        public int XfId { get; set; }
        #region Internal Properties
        int _numFmtId;
        internal int NumberFormatId
        {
            get
            {
                return this._numFmtId;
            }
            set
            {
                this._numFmtId = value;
                this.ApplyNumberFormat = (value>0);
            }
        }

        internal int FontId { get; set; }
        internal int FillId { get; set; }
        internal int BorderId { get; set; }
        private bool isBuildIn
        {
            get;
            set;
        }
        internal bool? ApplyNumberFormat
        {
            get;
            set;
        }
        internal bool? ApplyFont
        {
            get;
            set;
        }
        internal bool? ApplyFill
        {
            get;
            set;
        }
        internal bool? ApplyBorder
        {
            get;
            set;
        }
        internal bool? ApplyAlignment
        {
            get;
            set;
        } = true;
        internal bool? ApplyProtection
        {
            get;
            set;
        }
        #endregion
        #region Public Properties

        /// <summary>
        /// Numberformat properties
        /// </summary>
        public ExcelNumberFormatXml Numberformat 
        {
            get
            {
                return this._styles.NumberFormats[this._numFmtId < 0 ? 0 : this._numFmtId];
            }
        }
        /// <summary>
        /// Font properties
        /// </summary>
        public ExcelFontXml Font 
        { 
           get
           {
               return this._styles.Fonts[this.FontId < 0 ? 0 : this.FontId];
           }
        }
        /// <summary>
        /// Fill properties
        /// </summary>
        public ExcelFillXml Fill
        {
            get
            {
                return this._styles.Fills[this.FillId < 0 ? 0 : this.FillId];
            }
        }        
        /// <summary>
        /// Border style properties
        /// </summary>
        public ExcelBorderXml Border
        {
            get
            {
                return this._styles.Borders[this.BorderId < 0 ? 0 : this.BorderId];
            }
        }
        const string horizontalAlignPath = "d:alignment/@horizontal";

        /// <summary>
        /// Horizontal alignment
        /// </summary>
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; } = ExcelHorizontalAlignment.General;
        const string verticalAlignPath = "d:alignment/@vertical";

        /// <summary>
        /// Vertical alignment
        /// </summary>
        public ExcelVerticalAlignment VerticalAlignment { get; set; } = ExcelVerticalAlignment.Bottom;
        const string justifyLastLine = "d:alignment/@justifyLastLine";
        /// <summary>
        /// If the cells justified or distributed alignment should be used on the last line of text
        /// </summary>
        public bool JustifyLastLine { get; set; } = false;
        const string wrapTextPath = "d:alignment/@wrapText";

        /// <summary>
        /// Wraped text
        /// </summary>
        public bool WrapText { get; set; } = false;
        string textRotationPath = "d:alignment/@textRotation";
        int _textRotation = 0;
        /// <summary>
        /// Text rotation angle
        /// </summary>
        public int TextRotation
        {
            get
            {
                return (this._textRotation == int.MinValue ? 0 : this._textRotation);
            }
            set
            {
                this._textRotation = value;
            }
        }
        const string lockedPath = "d:protection/@locked";

        /// <summary>
        /// Locked when sheet is protected
        /// </summary>
        public bool Locked { get; set; } = true;
        const string hiddenPath = "d:protection/@hidden";

        /// <summary>
        /// Hide formulas when sheet is protected
        /// </summary>
        public bool Hidden { get; set; } = false;
        const string quotePrefixPath = "@quotePrefix";
        /// <summary>
        /// Prefix the formula with a quote.
        /// </summary>
        public bool QuotePrefix{ get; set; } = false;
        const string readingOrderPath = "d:alignment/@readingOrder";
        ExcelReadingOrder _readingOrder = ExcelReadingOrder.ContextDependent;
        /// <summary>
        /// Readingorder
        /// </summary>
        public ExcelReadingOrder ReadingOrder
        {
            get
            {
                return this._readingOrder;
            }
            set
            {
                this._readingOrder = value;
            }
        }
        const string shrinkToFitPath = "d:alignment/@shrinkToFit";

        /// <summary>
        /// Shrink to fit
        /// </summary>
        public bool ShrinkToFit { get; set; } = false;
        const string indentPath = "d:alignment/@indent";
        int _indent = 0;
        /// <summary>
        /// Indentation
        /// </summary>
        public int Indent
        {
            get
            {
                return (this._indent == int.MinValue ? 0 : this._indent);
            }
            set
            {
                this._indent=value;
            }
        }
        #endregion
        internal static void RegisterEvent(ExcelXfs xf)
        {
            //                RegisterEvent(xf, xf.Xf_ChangedEvent);
        }
        internal override string Id
        {

            get
            {
                return this.XfId + "|" + this.NumberFormatId.ToString() + "|" + this.FontId.ToString() + "|" + this.FillId.ToString() + "|" + this.BorderId.ToString() + this.VerticalAlignment.ToString() + "|" + this.HorizontalAlignment.ToString() + "|" + this.WrapText.ToString() + "|" + this.ReadingOrder.ToString() + "|" + this.isBuildIn.ToString() + this.TextRotation.ToString() + this.Locked.ToString() + this.Hidden.ToString() + this.ShrinkToFit.ToString() + this.Indent.ToString() + this.QuotePrefix.ToString() + this.JustifyLastLine.ToString(); 
                //return Numberformat.Id + "|" + Font.Id + "|" + Fill.Id + "|" + Border.Id + VerticalAlignment.ToString() + "|" + HorizontalAlignment.ToString() + "|" + WrapText.ToString() + "|" + ReadingOrder.ToString(); 
            }
        }
        internal ExcelXfs Copy()
        {
            return this.Copy(this._styles);
        }        
        internal ExcelXfs Copy(ExcelStyles styles)
        {
            ExcelXfs newXF = new ExcelXfs(this.NameSpaceManager, styles);
            newXF.NumberFormatId = this._numFmtId;
            newXF.FontId = this.FontId;
            newXF.FillId = this.FillId;
            newXF.BorderId = this.BorderId;
            newXF.XfId = this.XfId;
            newXF.ReadingOrder = this._readingOrder;
            newXF.HorizontalAlignment = this.HorizontalAlignment;
            newXF.VerticalAlignment = this.VerticalAlignment;
            newXF.WrapText = this.WrapText;
            newXF.ShrinkToFit = this.ShrinkToFit;
            newXF.Indent = this._indent;
            newXF.TextRotation = this._textRotation;
            newXF.Locked = this.Locked;
            newXF.Hidden = this.Hidden;
            newXF.QuotePrefix = this.QuotePrefix;
            newXF.JustifyLastLine = this.JustifyLastLine;
            newXF.ApplyAlignment = this.ApplyAlignment;
            newXF.ApplyBorder = this.ApplyBorder;
            newXF.ApplyFill = this.ApplyFill;
            newXF.ApplyFont = this.ApplyFont;
            newXF.ApplyNumberFormat= this.ApplyNumberFormat;
            newXF.ApplyProtection = this.ApplyProtection;            
            return newXF;
        }

        internal int GetNewID(ExcelStyleCollection<ExcelXfs> xfsCol, StyleBase styleObject, eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            ExcelXfs newXfs = this.Copy();
            switch(styleClass)
            {
                case eStyleClass.Numberformat:
                    newXfs.NumberFormatId = this.GetIdNumberFormat(styleProperty, value);
                    styleObject.SetIndex(newXfs.NumberFormatId);
                    break;
                case eStyleClass.Font:
                {
                    newXfs.FontId = this.GetIdFont(styleProperty, value);
                    styleObject.SetIndex(newXfs.FontId);
                    break;
                }
                case eStyleClass.Fill:
                case eStyleClass.FillBackgroundColor:
                case eStyleClass.FillPatternColor:
                    newXfs.FillId = this.GetIdFill(styleClass, styleProperty, value);
                    styleObject.SetIndex(newXfs.FillId);
                    break;
                case eStyleClass.GradientFill:
                case eStyleClass.FillGradientColor1:
                case eStyleClass.FillGradientColor2:
                    newXfs.FillId = this.GetIdGradientFill(styleClass, styleProperty, value);
                    styleObject.SetIndex(newXfs.FillId);
                    break;
                case eStyleClass.Border:
                case eStyleClass.BorderBottom:
                case eStyleClass.BorderDiagonal:
                case eStyleClass.BorderLeft:
                case eStyleClass.BorderRight:
                case eStyleClass.BorderTop:
                    newXfs.BorderId = this.GetIdBorder(styleClass, styleProperty, value);
                    styleObject.SetIndex(newXfs.BorderId);
                    break;
                case eStyleClass.Style:
                    switch(styleProperty)
                    {
                        case eStyleProperty.XfId:
                            newXfs.XfId = (int)value;
                            break;
                        case eStyleProperty.HorizontalAlign:
                            newXfs.HorizontalAlignment=(ExcelHorizontalAlignment)value;
                            break;
                        case eStyleProperty.VerticalAlign:
                            newXfs.VerticalAlignment = (ExcelVerticalAlignment)value;
                            break;
                        case eStyleProperty.WrapText:
                            newXfs.WrapText = (bool)value;
                            break;
                        case eStyleProperty.ReadingOrder:
                            newXfs.ReadingOrder = (ExcelReadingOrder)value;
                            break;
                        case eStyleProperty.ShrinkToFit:
                            newXfs.ShrinkToFit=(bool)value;
                            break;
                        case eStyleProperty.Indent:
                            newXfs.Indent = (int)value;
                            break;
                        case eStyleProperty.TextRotation:
                            newXfs.TextRotation = (int)value;
                            break;
                        case eStyleProperty.Locked:
                            newXfs.Locked = (bool)value;
                            break;
                        case eStyleProperty.Hidden:
                            newXfs.Hidden = (bool)value;
                            break;
                        case eStyleProperty.QuotePrefix:
                            newXfs.QuotePrefix = (bool)value;
                            break;
                        case eStyleProperty.JustifyLastLine:
                            newXfs.JustifyLastLine = (bool)value;
                            break;
                        default:
                            throw (new Exception("Invalid property for class style."));

                    }
                    break;
                default:
                    break;
            }
            int id = xfsCol.FindIndexById(newXfs.Id);
            if (id < 0)
            {
                return xfsCol.Add(newXfs.Id, newXfs);
            }
            return id;
        }

        private int GetIdBorder(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            ExcelBorderXml border = this.Border.Copy();

            switch (styleClass)
            {
                case eStyleClass.BorderBottom:
                    SetBorderItem(border.Bottom, styleProperty, value);
                    break;
                case eStyleClass.BorderDiagonal:
                    SetBorderItem(border.Diagonal, styleProperty, value);
                    break;
                case eStyleClass.BorderLeft:
                    SetBorderItem(border.Left, styleProperty, value);
                    break;
                case eStyleClass.BorderRight:
                    SetBorderItem(border.Right, styleProperty, value);
                    break;
                case eStyleClass.BorderTop:
                    SetBorderItem(border.Top, styleProperty, value);
                    break;
                case eStyleClass.Border:
                    if (styleProperty == eStyleProperty.BorderDiagonalUp)
                    {
                        border.DiagonalUp = (bool)value;
                    }
                    else if (styleProperty == eStyleProperty.BorderDiagonalDown)
                    {
                        border.DiagonalDown = (bool)value;
                    }
                    else
                    {
                        throw (new Exception("Invalid property for class Border."));
                    }
                    break;
                default:
                    throw (new Exception("Invalid class/property for class Border."));
            }

            string id = border.Id;
            int subId = this._styles.Borders.FindIndexById(id);
            if (subId == int.MinValue)
            {
                return this._styles.Borders.Add(id, border);
            }
            return subId;
        }

        private static void SetBorderItem(ExcelBorderItemXml excelBorderItem, eStyleProperty styleProperty, object value)
        {
            if(styleProperty==eStyleProperty.Style)
            {
                excelBorderItem.Style = (ExcelBorderStyle)value;
                return;
            }

            //Check that we have an style
            if (excelBorderItem.Style == ExcelBorderStyle.None)
            {
                throw (new InvalidOperationException("Can't set bordercolor when style is not set."));
            }

            if (styleProperty == eStyleProperty.Color)
            {
                excelBorderItem.Color.Rgb = value.ToString();
            }
            else if(styleProperty == eStyleProperty.Theme)
            {
                excelBorderItem.Color.Theme = (eThemeSchemeColor?)value;
            }
            else if (styleProperty == eStyleProperty.IndexedColor)
            {
                excelBorderItem.Color.Indexed = (int)value;
            }
            else if (styleProperty == eStyleProperty.Tint)
            {
                excelBorderItem.Color.Tint = (decimal)value;
            }
            else if (styleProperty == eStyleProperty.AutoColor)
            {
                excelBorderItem.Color.Auto = (bool)value;
            }
        }

        private int GetIdFill(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            ExcelFillXml fill = this.Fill.Copy();

            switch (styleProperty)
            {
                case eStyleProperty.PatternType:
                    if (fill is ExcelGradientFillXml)
                    {
                        fill = new ExcelFillXml(this.NameSpaceManager);
                    }
                    fill.PatternType = (ExcelFillStyle)value;
                    break;
                case eStyleProperty.Color:
                case eStyleProperty.Tint:
                case eStyleProperty.IndexedColor:
                case eStyleProperty.AutoColor:
                case eStyleProperty.Theme:
                    if (fill is ExcelGradientFillXml)
                    {
                        fill = new ExcelFillXml(this.NameSpaceManager);
                    }
                    if (fill.PatternType == ExcelFillStyle.None)
                    {
                        throw (new ArgumentException("Can't set color when patterntype is not set."));
                    }
                    ExcelColorXml destColor;
                    if (styleClass==eStyleClass.FillPatternColor)
                    {
                        destColor = fill.PatternColor;
                    }
                    else
                    {
                        destColor = fill.BackgroundColor;
                    }

                    if (styleProperty == eStyleProperty.Color)
                    {
                        destColor.Rgb = value.ToString();
                    }
                    else if (styleProperty == eStyleProperty.Tint)
                    {
                        destColor.Tint = (decimal)value;
                    }
                    else if (styleProperty == eStyleProperty.IndexedColor)
                    {
                        destColor.Indexed = (int)value;
                    }
                    else if(styleProperty == eStyleProperty.Theme)
                    {
                        destColor.Theme = (eThemeSchemeColor?)value;
                    }
                    else
                    {
                        destColor.Auto = (bool)value;
                    }

                    break;
                default:
                    throw (new ArgumentException("Invalid class/property for class Fill."));
            }

            string id = fill.Id;
            int subId = this._styles.Fills.FindIndexById(id);
            if (subId == int.MinValue)
            {
                return this._styles.Fills.Add(id, fill);
            }
            return subId;
        }
        private int GetIdGradientFill(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            ExcelGradientFillXml fill;
            if(this.Fill is ExcelGradientFillXml)
            {
                fill = (ExcelGradientFillXml)this.Fill.Copy();
            }
            else
            {
                fill = new ExcelGradientFillXml(this.Fill.NameSpaceManager);
                fill.GradientColor1.SetColor(Color.White);
                fill.GradientColor2.SetColor(Color.FromArgb(79,129,189));
                fill.Type=ExcelFillGradientType.Linear;
                fill.Degree=90;
                fill.Top = double.NaN;
                fill.Bottom = double.NaN;
                fill.Left = double.NaN;
                fill.Right = double.NaN;
            }

            switch (styleProperty)
            {
                case eStyleProperty.GradientType:
                    fill.Type = (ExcelFillGradientType)value;
                    break;
                case eStyleProperty.GradientDegree:
                    fill.Degree = (double)value;
                    break;
                case eStyleProperty.GradientTop:
                    fill.Top = (double)value;
                    break;
                case eStyleProperty.GradientBottom: 
                    fill.Bottom = (double)value;
                    break;
                case eStyleProperty.GradientLeft:
                    fill.Left = (double)value;
                    break;
                case eStyleProperty.GradientRight:
                    fill.Right = (double)value;
                    break;
                case eStyleProperty.Color:
                case eStyleProperty.Tint:
                case eStyleProperty.IndexedColor:
                case eStyleProperty.AutoColor:
                case eStyleProperty.Theme:
                    ExcelColorXml destColor;

                    if (styleClass == eStyleClass.FillGradientColor1)
                    {
                        destColor = fill.GradientColor1;
                    }
                    else
                    {
                        destColor = fill.GradientColor2;
                    }
                    
                    if (styleProperty == eStyleProperty.Color)
                    {
                        destColor.Rgb = value.ToString();
                    }
                    else if (styleProperty == eStyleProperty.Tint)
                    {
                        destColor.Tint = (decimal)value;
                    }
                    else if (styleProperty == eStyleProperty.Theme)
                    {
                        destColor.Theme = (eThemeSchemeColor?)value;
                    }
                    else if (styleProperty == eStyleProperty.IndexedColor)
                    {
                        destColor.Indexed = (int)value;
                    }
                    else
                    {
                        destColor.Auto = (bool)value;
                    }
                    break;
                default:
                    throw (new ArgumentException("Invalid class/property for class Fill."));
            }

            string id = fill.Id;
            int subId = this._styles.Fills.FindIndexById(id);
            if (subId == int.MinValue)
            {
                return this._styles.Fills.Add(id, fill);
            }
            return subId;
        }

        private int GetIdNumberFormat(eStyleProperty styleProperty, object value)
        {
            if (styleProperty == eStyleProperty.Format)
            {
                ExcelNumberFormatXml item=null;
                if (!this._styles.NumberFormats.FindById(value.ToString(), ref item))
                {
                    item = new ExcelNumberFormatXml(this.NameSpaceManager) { Format = value.ToString(), NumFmtId = this._styles.NumberFormats.NextId++ };
                    this._styles.NumberFormats.Add(value.ToString(), item);
                }
                return item.NumFmtId;
            }
            else
            {
                throw (new Exception("Invalid property for class Numberformat"));
            }
        }
        private int GetIdFont(eStyleProperty styleProperty, object value)
        {
            ExcelFontXml fnt = this.Font.Copy();

            switch (styleProperty)
            {
                case eStyleProperty.Name:
                    fnt.Name = value.ToString();
                    break;
                case eStyleProperty.Size:
                    fnt.Size = (float)value;
                    break;
                case eStyleProperty.Family:
                    fnt.Family = (int)value;
                    break;
                case eStyleProperty.Bold:
                    fnt.Bold = (bool)value;
                    break;
                case eStyleProperty.Italic:
                    fnt.Italic = (bool)value;
                    break;
                case eStyleProperty.Strike:
                    fnt.Strike = (bool)value;
                    break;
                case eStyleProperty.UnderlineType:
                    fnt.UnderLineType = (ExcelUnderLineType)value;
                    break;
                case eStyleProperty.Color:
                    fnt.Color.Rgb=value.ToString();
                    break;
                case eStyleProperty.Tint:
                    fnt.Color.Tint = (decimal)value;
                    break;
                case eStyleProperty.Theme:
                    fnt.Color.Theme = (eThemeSchemeColor?)value;
                    break;
                case eStyleProperty.IndexedColor:
                    fnt.Color.Indexed = (int)value;
                    break;
                case eStyleProperty.AutoColor:
                    fnt.Color.Auto = (bool)value;
                    break;
                case eStyleProperty.VerticalAlign:
                    fnt.VerticalAlign = ((ExcelVerticalAlignmentFont)value) == ExcelVerticalAlignmentFont.None ? "" : value.ToString().ToLower(CultureInfo.InvariantCulture);
                    break;
                case eStyleProperty.Scheme:
                    fnt.Scheme = value.ToString();
                    break;
                case eStyleProperty.Charset:
                    fnt.Charset = (int?)value;
                    break;
                default:
                    throw (new Exception("Invalid property for class Font"));
            }

            string id = fnt.Id;
            int subId = this._styles.Fonts.FindIndexById(id);
            if (subId == int.MinValue)
            {
                return this._styles.Fonts.Add(id,fnt);
            }
            return subId;
        }
        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            return this.CreateXmlNode(topNode, false);
        }
        internal XmlNode CreateXmlNode(XmlNode topNode, bool isCellStyleXsf)
        {
            this.TopNode = topNode;
            if(this.XfId<0 || this.XfId>= this._styles.CellStyleXfs.Count) //XfId has an invalid reference. Remove it.
            {
                this.XfId = int.MinValue;
            }

            bool doSetXfId = (!isCellStyleXsf && this.XfId > int.MinValue && this._styles.CellStyleXfs.Count > 0 && this._styles.CellStyleXfs[this.XfId].newID >= 0);
            if (this._numFmtId >= 0)
            {
                this.SetXmlNodeString("@numFmtId", this._numFmtId.ToString());
                if(this._numFmtId > 0 || this.ApplyNumberFormat.HasValue)
                {
                    this.SetXmlNodeBool("@applyNumberFormat", this.ApplyNumberFormat??true);
                }
            }
            if (this.FontId >= 0)
            {
                this.SetXmlNodeString("@fontId", this._styles.Fonts[this.FontId].newID.ToString());
                if (this.FontId > 0 || this.ApplyFont.HasValue)
                {
                    this.SetXmlNodeBool("@applyFont", this.ApplyFont ?? true);
                }
            }
            if (this.FillId >= 0)
            {
                this.SetXmlNodeString("@fillId", this._styles.Fills[this.FillId].newID.ToString());
                if (this.FillId > 0 || this.ApplyFill.HasValue)
                {
                    this.SetXmlNodeBool("@applyFill", this.ApplyFill ?? true);
                }
            }
            if (this.BorderId >= 0)
            {
                this.SetXmlNodeString("@borderId", this._styles.Borders[this.BorderId].newID.ToString());
                if (this.BorderId > 0 || this.ApplyBorder.HasValue)
                {
                    this.SetXmlNodeBool("@applyBorder", this.ApplyBorder ?? true);
                }
            }
            if(this.HorizontalAlignment != ExcelHorizontalAlignment.General)
            {
                this.SetXmlNodeString(horizontalAlignPath, SetAlignString(this.HorizontalAlignment));
            }

            if (doSetXfId)
            {
                this.SetXmlNodeString("@xfId", this._styles.CellStyleXfs[this.XfId].newID.ToString());
            }

            if(this.VerticalAlignment != ExcelVerticalAlignment.Bottom)
            {
                this.SetXmlNodeString(verticalAlignPath, SetAlignString(this.VerticalAlignment));
            }

            if(this.WrapText)
            {
                this.SetXmlNodeString(wrapTextPath, "1");
            }

            if(this._readingOrder!=ExcelReadingOrder.ContextDependent)
            {
                this.SetXmlNodeString(readingOrderPath, ((int)this._readingOrder).ToString());
            }

            if(this.ShrinkToFit)
            {
                this.SetXmlNodeString(shrinkToFitPath, "1");
            }

            if(this._indent > 0)
            {
                this.SetXmlNodeString(indentPath, this._indent.ToString());
            }

            if(this._textRotation > 0)
            {
                this.SetXmlNodeString(this.textRotationPath, this._textRotation.ToString());
            }

            if(!this.Locked)
            {
                this.SetXmlNodeString(lockedPath, "0");
            }

            if(this.Hidden)
            {
                this.SetXmlNodeString(hiddenPath, "1");
            }

            if(this.QuotePrefix)
            {
                this.SetXmlNodeString(quotePrefixPath, "1");
            }

            if(this.JustifyLastLine)
            {
                this.SetXmlNodeString(justifyLastLine, "1");
            }

            if ((this.Locked == false || this.Hidden == true || this.ApplyProtection.HasValue)) //Not default values, apply protection.
            {
                this.SetXmlNodeBool("@applyProtection", this.ApplyProtection??true);
            }

            if (this.HorizontalAlignment != ExcelHorizontalAlignment.General || this.VerticalAlignment != ExcelVerticalAlignment.Bottom || this.ApplyProtection.HasValue)
            {
                this.SetXmlNodeBool("@applyAlignment", this.ApplyProtection??true);
            }

            return this.TopNode;
        }

        private static string SetAlignString(Enum align)
        {
            string newName = Enum.GetName(align.GetType(), align);
            return newName.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + newName.Substring(1, newName.Length - 1);
        }
    }
}
