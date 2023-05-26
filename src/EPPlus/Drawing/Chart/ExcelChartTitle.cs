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
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// The title of a chart
    /// </summary>
    public abstract class ExcelChartTitle : XmlHelper, IDrawingStyle, IStyleMandatoryProperties
    {
        internal ExcelChart _chart;
        internal string _nsPrefix = "";
        private readonly string titlePath = "{0}:tx/{0}:rich/a:p/a:r/a:t";

        internal ExcelChartTitle(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode node, string nsPrefix) :
            base(nameSpaceManager, node)
        {
            this._chart = chart;
            this._nsPrefix = nsPrefix;
            this.titlePath = string.Format(this.titlePath, nsPrefix);
            if (chart._isChartEx)
            {
                this.AddSchemaNodeOrder(new string[] { "tx", "strRef", "rich", "bodyPr", "lstStyle", "layout", "p", "overlay", "spPr", "txPr" }, ExcelDrawing._schemaNodeOrderSpPr);
                this.CreateTopNode();
            }
            else
            {
                this.AddSchemaNodeOrder(this._chart._chartXmlHelper.SchemaNodeOrder, ExcelDrawing._schemaNodeOrderSpPr);
                this.CreateTopNode();
                if (this.TopNode.HasChildNodes == false)
                {
                    this.TopNode.InnerXml = GetInitXml("c");
                    chart.ApplyStyleOnPart(this, chart.StyleManager?.Style?.Title, true);
                }
            }

        }

        private void CreateTopNode()
        {            
            if (this.TopNode.LocalName != "title")
            {
                this.TopNode = this.CreateNode(this._nsPrefix+":title");
            }
        }

        internal static string GetInitXml(string prefix)
        {
            return $"<{prefix}:tx><{prefix}:rich><a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\" />" +
                    $"<a:lstStyle />" +
                    $"<a:p><a:pPr>" +
                    $"<a:defRPr sz=\"1080\" b=\"1\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">" +
                    "<a:effectLst/><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr>" +
                    $"</a:pPr><a:r><a:t/></a:r></a:p></{prefix}:rich></{prefix}:tx><{prefix}:layout /><{prefix}:overlay val=\"0\" />" +
                    $"<{prefix}:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></{prefix}:spPr>";
        }

        /// <summary>
        /// The text
        /// </summary>
        public abstract string Text
        {
            get;
            set;
        }
                ExcelDrawingBorder _border = null;
        /// <summary>
        /// A reference to the border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (this._border == null)
                {
                    this._border = new ExcelDrawingBorder(this._chart, this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr/a:ln", this.SchemaNodeOrder);
                }
                return this._border;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (this._fill == null)
                {
                    this._fill = new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr", this.SchemaNodeOrder);
                }
                return this._fill;
            }
        }
        ExcelTextFont _font=null;
        /// <summary>
        /// A reference to the font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (this._font == null)
                {
                    if (this._richText == null || this._richText.Count == 0)
                    {
                        this.RichText.Add("");
                    }

                    this._font = new ExcelTextFont(this._chart, this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:p/a:pPr/a:defRPr", this.SchemaNodeOrder);
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
                    this._textBody = new ExcelTextBody(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr", this.SchemaNodeOrder);
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
                    this._effect = new ExcelDrawingEffectStyle(this._chart, this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr/a:effectLst", this.SchemaNodeOrder);
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
                    this._threeD = new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr", this.SchemaNodeOrder);
                }
                return this._threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            this.CreatespPrNode($"{this._nsPrefix}:spPr");
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
                    float defFont = 14;
                    ExcelChartStyleEntry? stylePart = this.GetStylePart();
                    if(stylePart!=null && stylePart.HasTextRun)
                    {
                        defFont = Convert.ToSingle(stylePart.DefaultTextRun.FontSize);
                    }

                    this._richText = new ExcelParagraphCollection(this._chart, this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:tx/{this._nsPrefix }:rich/a:p", this.SchemaNodeOrder, defFont);
                }
                return this._richText;
            }
        }

        private ExcelChartStyleEntry GetStylePart()
        {
            ExcelChartStyle? style = this._chart._styleManager?.Style;
            if (style == null)
            {
                return null;
            }

            if (this.TopNode.ParentNode.LocalName == "chart")
            {
                return this._chart._styleManager.Style.Title;
            }
            else
            {
                return this._chart._styleManager.Style.AxisTitle;
            }
        }

        /// <summary>
        /// Show without overlaping the chart.
        /// </summary>
        public bool Overlay
        {
            get
            {
                if (this._chart._isChartEx)
                {
                    return this.GetXmlNodeBool("@overlay");
                }
                else
                {
                    return this.GetXmlNodeBool("c:overlay/@val");
                }
            }
            set
            {
                if (this._chart._isChartEx)
                {
                    this.SetXmlNodeBool("@overlay", value);
                }
                else
                {
                    this.SetXmlNodeBool("c:overlay/@val", value);
                }
            }
        }
        /// <summary>
        /// The centering of the text. Centers the text to the smallest possible text container.
        /// </summary>
        public bool AnchorCtr
        {
            get
            {
                return this.GetXmlNodeBool($"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr/@anchorCtr", false);
            }
            set
            {
                this.SetXmlNodeBool($"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr/@anchorCtr", value, false);
            }
        }
        /// <summary>
        /// How the text is anchored
        /// </summary>
        public eTextAnchoringType Anchor
        {
            get
            {
                return this.GetXmlNodeString($"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr/@anchor").TranslateTextAchoring();
            }
            set
            {
                this.SetXmlNodeString($"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr/@anchorCtr", value.TranslateTextAchoringText());
            }
        }
        const string TextVerticalPath = "xdr:sp/xdr:txBody/a:bodyPr/@vert";
        /// <summary>
        /// Vertical text
        /// </summary>
        public eTextVerticalType TextVertical
        {
            get
            {
                return this.GetXmlNodeString($"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr/@vert").TranslateTextVertical();
            }
            set
            {
                this.SetXmlNodeString($"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr/@vert", value.TranslateTextVerticalText());
            }
        }
        /// <summary>
        /// Rotation in degrees (0-360)
        /// </summary>
        public double Rotation
        {
            get
            {
                int i= this.GetXmlNodeInt($"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr/@rot");
                if (i < 0)
                {
                    return 360 - (i / 60000);
                }
                else
                {
                    return (i / 60000);
                }
            }
            set
            {
                int v;
                if(value <0 || value > 360)
                {
                    throw(new ArgumentOutOfRangeException("Rotation must be between 0 and 360"));
                }

                if (value > 180)
                {
                    v = (int)((value - 360) * 60000);
                }
                else
                {
                    v = (int)(value * 60000);
                }

                this.SetXmlNodeString($"{this._nsPrefix}:tx/{this._nsPrefix}:rich/a:bodyPr/@rot", v.ToString());
            }
        }

        void IStyleMandatoryProperties.SetMandatoryProperties()
        {
            this.TextBody.Anchor = eTextAnchoringType.Center;
            this.TextBody.AnchorCenter = true;
            this.TextBody.WrapText = eTextWrappingType.Square;
            this.TextBody.VerticalTextOverflow = eTextVerticalOverflow.Ellipsis;
            this.TextBody.ParagraphSpacing = true;
            this.TextBody.Rotation = 0;

            if (this.Font.Kerning == 0)
            {
                this.Font.Kerning = 12;
            }

            this.Font.Bold = this.Font.Bold; //Must be set

            this.CreatespPrNode($"{this._nsPrefix}:spPr");
        }
    }
    public class ExcelChartTitleStandard : ExcelChartTitle
    {
        internal ExcelChartTitleStandard(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode node, string nsPrefix) : base(chart, nameSpaceManager, node, nsPrefix)
        {
            this.titleLinkPath = string.Format(this.titleLinkPath, nsPrefix);
        }
        private readonly string titleLinkPath = "{0}:tx/{0}:strRef";
        public override string Text 
        {
            get
            {
                if (this.LinkedCell == null)
                {
                    return this.RichText.Text;
                }
                else
                {
                    return this.LinkedCell.Text;
                }
            }
            set
            {
                bool applyStyle = (this.RichText.Count == 0);
                this.LinkedCell = null;
                this.RichText.Text = value;
                if (applyStyle)
                {
                    this._chart.ApplyStyleOnPart(this, this._chart.StyleManager?.Style?.Title, true);
                }
            }
        }
        /// <summary>
        /// A reference to a cell used as the title text
        /// </summary>
        public ExcelRangeBase LinkedCell
        {
            get
            {
                string? a = this.GetXmlNodeString($"{this.titleLinkPath}/c:f");
                if (ExcelCellBase.IsValidAddress(a))
                {
                    ExcelAddressBase? address = new ExcelAddressBase(a);
                    ExcelWorksheet ws;
                    if (string.IsNullOrEmpty(address.WorkSheetName))
                    {
                        ws = this._chart.WorkSheet;
                    }
                    else
                    {
                        ws = this._chart.WorkSheet.Workbook.Worksheets[address.WorkSheetName];
                    }
                    if (ws == null)
                    {
                        return null;
                    }

                    return ws.Cells[address.LocalAddress];
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    this.DeleteNode($"{this._nsPrefix}:tx/{this._nsPrefix}:strRef");
                    this.RichText.Text = "";
                    return;
                }
                else
                {
                    this.DeleteNode($"{this._nsPrefix}:tx/{this._nsPrefix}:rich");
                    this.SetXmlNodeString($"{this.titleLinkPath}/c:f", value.FullAddressAbsolute);
                    XmlNode? cache = this.CreateNode($"{this._nsPrefix}:tx/{this._nsPrefix}strRef/c:strCache", false, true);
                    cache.InnerXml = $"<c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{value.Text}</c:v></c:pt>";
                }
            }
        }
    }
}
