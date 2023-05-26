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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Access to trendline label properties
    /// </summary>
    public class ExcelChartTrendlineLabel : XmlHelper, IDrawingStyle
    {
        ExcelChartStandardSerie _serie;        
        internal ExcelChartTrendlineLabel(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelChartStandardSerie serie) : base(namespaceManager, topNode)
        {
            this._serie = serie;

            this.AddSchemaNodeOrder(new string[] { "layout", "tx", "numFmt", "spPr", "txPr" }, ExcelDrawing._schemaNodeOrderSpPr);
        }

        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (this._fill == null)
                {
                    this._fill = new ExcelDrawingFill(this._serie._chart, this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:spPr", this.SchemaNodeOrder);
                }
                return this._fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Access to border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (this._border == null)
                {
                    this._border = new ExcelDrawingBorder(this._serie._chart, this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:spPr/a:ln", this.SchemaNodeOrder);
                }
                return this._border;
            }
        }
        ExcelTextFont _font = null;
        /// <summary>
        /// Access to font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (this._font == null)
                {
                    this._font = new ExcelTextFont(this._serie._chart, this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:txPr/a:p/a:pPr/a:defRPr", this.SchemaNodeOrder);
                }
                return this._font;
            }
        }
        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public ExcelTextBody TextBody
        {
            get
            {
                if (this._textBody == null)
                {
                    this._textBody = new ExcelTextBody(this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:txPr/a:bodyPr", this.SchemaNodeOrder);
                }
                return this._textBody;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (this._effect == null)
                {
                    this._effect = new ExcelDrawingEffectStyle(this._serie._chart, this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:spPr/a:effectLst", this.SchemaNodeOrder);
                }
                return this._effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (this._threeD == null)
                {
                    this._threeD = new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:spPr", this.SchemaNodeOrder);
                }
                return this._threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            this.CreatespPrNode("c:trendlineLbl/c:spPr");
        }

        ExcelParagraphCollection _richText = null;
        /// <summary>
        /// Richtext
        /// </summary>
        public ExcelParagraphCollection RichText
        {
            get
            {
                if (this._richText == null)
                {
                    this._richText = new ExcelParagraphCollection(this._serie._chart, this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:tx/c:rich/a:p", this.SchemaNodeOrder);
                }
                return this._richText;
            }
        }
        /// <summary>
        /// Numberformat
        /// </summary>
        public string NumberFormat
        {
            get
            {
                return this.GetXmlNodeString("c:trendlineLbl/c:numFmt/@formatCode");
            }
            set
            {
                this.SetXmlNodeString("c:trendlineLbl/c:numFmt/@formatCode", value);
            }
        }
        /// <summary>
        /// If the numberformat is linked to the source data
        /// </summary>
        public bool SourceLinked
        {
            get
            {
                return this.GetXmlNodeBool("c:trendlineLbl/c:numFmt/@sourceLinked");
            }
            set
            {
                this.SetXmlNodeBool("c:trendlineLbl/c:numFmt/@sourceLinked", value, true);
            }
        }        
    }
}