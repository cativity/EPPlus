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

using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Data table on chart level. 
/// </summary>
public class ExcelChartDataTable : XmlHelper, IDrawingStyle
{
    ExcelChart _chart;

    internal ExcelChartDataTable(ExcelChart chart, XmlNamespaceManager ns, XmlNode node)
        : base(ns, node)
    {
        this.AddSchemaNodeOrder(new string[] { "dTable", "showHorzBorder", "showVertBorder", "showOutline", "showKeys", "spPr", "txPr" },
                                ExcelDrawing._schemaNodeOrderSpPr);

        XmlNode topNode = node.SelectSingleNode("c:dTable", this.NameSpaceManager);

        if (topNode == null)
        {
            topNode = node.OwnerDocument.CreateElement("c", "dTable", ExcelPackage.schemaChart);
            InserAfter(node, "c:valAx,c:catAx,c:dateAx,c:serAx", topNode);

            topNode.InnerXml = "<c:showHorzBorder val=\"1\"/><c:showVertBorder val=\"1\"/><c:showOutline val=\"1\"/><c:showKeys val=\"1\"/>"
                               + "<c:spPr><a:noFill/><a:ln cap = \"flat\" w=\"9525\" algn=\"ctr\" cmpd=\"sng\" ><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>"
                               + "<c:txPr><a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>"
                               + "<a:lstStyle/><a:p><a:pPr rtl=\"0\"><a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\"><a:solidFill><a:schemeClr val=\"dk1\"></a:schemeClr></a:solidFill>"
                               + "<a:latin typeface=\" + mn - lt\"/><a:ea typeface=\" + mn - ea\"/><a:cs typeface=\" + mn - cs\"/></a:defRPr></a:pPr><a:endParaRPr lang=\"en - US\"/></a:p></c:txPr>";
        }

        this.TopNode = topNode;
        this._chart = chart;
    }

    #region "Public properties"

    const string showHorzBorderPath = "c:showHorzBorder/@val";

    /// <summary>
    /// The horizontal borders will be shown in the data table
    /// </summary>
    public bool ShowHorizontalBorder
    {
        get => this.GetXmlNodeBool(showHorzBorderPath);
        set => this.SetXmlNodeString(showHorzBorderPath, value ? "1" : "0");
    }

    const string showVertBorderPath = "c:showVertBorder/@val";

    /// <summary>
    /// The vertical borders will be shown in the data table
    /// </summary>
    public bool ShowVerticalBorder
    {
        get => this.GetXmlNodeBool(showVertBorderPath);
        set => this.SetXmlNodeString(showVertBorderPath, value ? "1" : "0");
    }

    const string showOutlinePath = "c:showOutline/@val";

    /// <summary>
    /// The outline will be shown on the data table
    /// </summary>
    public bool ShowOutline
    {
        get => this.GetXmlNodeBool(showOutlinePath);
        set => this.SetXmlNodeString(showOutlinePath, value ? "1" : "0");
    }

    const string showKeysPath = "c:showKeys/@val";

    /// <summary>
    /// The legend keys will be shown in the data table
    /// </summary>
    public bool ShowKeys
    {
        get => this.GetXmlNodeBool(showKeysPath);
        set => this.SetXmlNodeString(showKeysPath, value ? "1" : "0");
    }

    ExcelDrawingFill _fill;

    /// <summary>
    /// Access fill properties
    /// </summary>
    public ExcelDrawingFill Fill => this._fill ??= new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);

    ExcelDrawingBorder _border;

    /// <summary>
    /// Access border properties
    /// </summary>
    public ExcelDrawingBorder Border => this._border ??= new ExcelDrawingBorder(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:ln", this.SchemaNodeOrder);

    ExcelTextFont _font;

    /// <summary>
    /// Access font properties
    /// </summary>
    public ExcelTextFont Font
    {
        get
        {
            if (this._font == null)
            {
                if (this.TopNode.SelectSingleNode("c:txPr", this.NameSpaceManager) == null)
                {
                    _ = this.CreateNode("c:txPr/a:bodyPr");
                    _ = this.CreateNode("c:txPr/a:lstStyle");
                }

                this._font = new ExcelTextFont(this._chart, this.NameSpaceManager, this.TopNode, "c:txPr/a:p/a:pPr/a:defRPr", this.SchemaNodeOrder);
            }

            return this._font;
        }
    }

    ExcelTextBody _textBody;

    /// <summary>
    /// Access to text body properties
    /// </summary>
    public ExcelTextBody TextBody => this._textBody ??= new ExcelTextBody(this.NameSpaceManager, this.TopNode, "c:txPr/a:bodyPr", this.SchemaNodeOrder);

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect => this._effect ??= new ExcelDrawingEffectStyle(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:effectLst", this.SchemaNodeOrder);

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD => this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);

    void IDrawingStyleBase.CreatespPr() => this.CreatespPrNode();

    #endregion
}