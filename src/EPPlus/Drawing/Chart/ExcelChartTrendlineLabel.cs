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

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Access to trendline label properties
/// </summary>
public class ExcelChartTrendlineLabel : XmlHelper, IDrawingStyle
{
    ExcelChartStandardSerie _serie;

    internal ExcelChartTrendlineLabel(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelChartStandardSerie serie)
        : base(namespaceManager, topNode)
    {
        this._serie = serie;

        this.AddSchemaNodeOrder(new string[] { "layout", "tx", "numFmt", "spPr", "txPr" }, ExcelDrawing._schemaNodeOrderSpPr);
    }

    ExcelDrawingFill _fill;

    /// <summary>
    /// Access to fill properties
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            return this._fill ??= new ExcelDrawingFill(this._serie._chart, this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:spPr", this.SchemaNodeOrder);
        }
    }

    ExcelDrawingBorder _border;

    /// <summary>
    /// Access to border properties
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get
        {
            return this._border ??= new ExcelDrawingBorder(this._serie._chart,
                                                           this.NameSpaceManager,
                                                           this.TopNode,
                                                           "c:trendlineLbl/c:spPr/a:ln",
                                                           this.SchemaNodeOrder);
        }
    }

    ExcelTextFont _font;

    /// <summary>
    /// Access to font properties
    /// </summary>
    public ExcelTextFont Font
    {
        get
        {
            return this._font ??= new ExcelTextFont(this._serie._chart,
                                                    this.NameSpaceManager,
                                                    this.TopNode,
                                                    "c:trendlineLbl/c:txPr/a:p/a:pPr/a:defRPr",
                                                    this.SchemaNodeOrder);
        }
    }

    ExcelTextBody _textBody;

    /// <summary>
    /// Access to text body properties
    /// </summary>
    public ExcelTextBody TextBody
    {
        get { return this._textBody ??= new ExcelTextBody(this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:txPr/a:bodyPr", this.SchemaNodeOrder); }
    }

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            return this._effect ??= new ExcelDrawingEffectStyle(this._serie._chart,
                                                                this.NameSpaceManager,
                                                                this.TopNode,
                                                                "c:trendlineLbl/c:spPr/a:effectLst",
                                                                this.SchemaNodeOrder);
        }
    }

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get { return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:trendlineLbl/c:spPr", this.SchemaNodeOrder); }
    }

    void IDrawingStyleBase.CreatespPr()
    {
        this.CreatespPrNode("c:trendlineLbl/c:spPr");
    }

    ExcelParagraphCollection _richText;

    /// <summary>
    /// Richtext
    /// </summary>
    public ExcelParagraphCollection RichText
    {
        get
        {
            return this._richText ??= new ExcelParagraphCollection(this._serie._chart,
                                                                   this.NameSpaceManager,
                                                                   this.TopNode,
                                                                   "c:trendlineLbl/c:tx/c:rich/a:p",
                                                                   this.SchemaNodeOrder);
        }
    }

    /// <summary>
    /// Numberformat
    /// </summary>
    public string NumberFormat
    {
        get { return this.GetXmlNodeString("c:trendlineLbl/c:numFmt/@formatCode"); }
        set { this.SetXmlNodeString("c:trendlineLbl/c:numFmt/@formatCode", value); }
    }

    /// <summary>
    /// If the numberformat is linked to the source data
    /// </summary>
    public bool SourceLinked
    {
        get { return this.GetXmlNodeBool("c:trendlineLbl/c:numFmt/@sourceLinked"); }
        set { this.SetXmlNodeBool("c:trendlineLbl/c:numFmt/@sourceLinked", value, true); }
    }
}