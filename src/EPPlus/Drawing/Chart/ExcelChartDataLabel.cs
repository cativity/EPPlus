/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/24/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Datalabel on chart level. 
/// This class is inherited by ExcelChartSerieDataLabel
/// </summary>
public abstract class ExcelChartDataLabel : XmlHelper, IDrawingStyle
{
    internal ExcelChart _chart;
    //internal string _nodeName;
    private string _nsPrefix;
    private readonly string _formatPath;
    private readonly string _sourceLinkedPath;

    internal ExcelChartDataLabel(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string nodeName, string nsPrefix)
        : base(ns, node)
    {
        //this._nodeName = nodeName;
        this._chart = chart;
        this._nsPrefix = nsPrefix;
        this._formatPath = $"{nsPrefix}:numFmt/@formatCode";
        this._sourceLinkedPath = $"{nsPrefix}:numFmt/@sourceLinked";
    }

    #region "Public properties"

    /// <summary>
    /// The position of the data labels
    /// </summary>
    public abstract eLabelPosition Position { get; set; }

    /// <summary>
    /// Show the values 
    /// </summary>
    public abstract bool ShowValue { get; set; }

    /// <summary>
    /// Show category names  
    /// </summary>
    public abstract bool ShowCategory { get; set; }

    /// <summary>
    /// Show series names
    /// </summary>
    public abstract bool ShowSeriesName { get; set; }

    /// <summary>
    /// Show percent values
    /// </summary>
    public abstract bool ShowPercent { get; set; }

    /// <summary>
    /// Show the leader lines
    /// </summary>
    public abstract bool ShowLeaderLines { get; set; }

    /// <summary>
    /// Show Bubble Size
    /// </summary>
    public abstract bool ShowBubbleSize { get; set; }

    /// <summary>
    /// Show the Lengend Key
    /// </summary>
    public abstract bool ShowLegendKey { get; set; }

    /// <summary>
    /// Separator string 
    /// </summary>
    public abstract string Separator { get; set; }

    /// <summary>
    /// The Numberformat string.
    /// </summary>
    public string Format
    {
        get => this.GetXmlNodeString(this._formatPath);
        set
        {
            this.SetXmlNodeString(this._formatPath, value);

            if (string.IsNullOrEmpty(value))
            {
                this.SourceLinked = true;
            }
            else
            {
                this.SourceLinked = false;
            }
        }
    }

    /// <summary>
    /// The Numberformats are linked to the source data.
    /// </summary>
    public bool SourceLinked
    {
        get => this.GetXmlNodeBool(this._sourceLinkedPath);
        set => this.SetXmlNodeBool(this._sourceLinkedPath, value);
    }

    ExcelDrawingFill _fill;

    /// <summary>
    /// Access fill properties
    /// </summary>
    public ExcelDrawingFill Fill => this._fill ??= new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr", this.SchemaNodeOrder);

    ExcelDrawingBorder _border;

    /// <summary>
    /// Access border properties
    /// </summary>
    public ExcelDrawingBorder Border =>
        this._border ??= new ExcelDrawingBorder(this._chart,
                                                this.NameSpaceManager,
                                                this.TopNode,
                                                $"{this._nsPrefix}:spPr/a:ln",
                                                this.SchemaNodeOrder);

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect =>
        this._effect ??= new ExcelDrawingEffectStyle(this._chart,
                                                     this.NameSpaceManager,
                                                     this.TopNode,
                                                     $"{this._nsPrefix}:spPr/a:effectLst",
                                                     this.SchemaNodeOrder);

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD => this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr", this.SchemaNodeOrder);

    ExcelTextFont _font;

    /// <summary>
    /// Access font properties
    /// </summary>
    public ExcelTextFont Font =>
        this._font ??= new ExcelTextFont(this._chart,
                                         this.NameSpaceManager,
                                         this.TopNode,
                                         $"{this._nsPrefix}:txPr/a:p/a:pPr/a:defRPr",
                                         this.SchemaNodeOrder,
                                         this.CreateDefaultText);

    void IDrawingStyleBase.CreatespPr() => this.CreatespPrNode();

    private void CreateDefaultText()
    {
        if (this.TopNode.SelectSingleNode($"{this._nsPrefix}:txPr", this.NameSpaceManager) == null)
        {
            if (!this.ExistsNode($"{this._nsPrefix}:spPr"))
            {
                XmlNode? spNode = this.CreateNode($"{this._nsPrefix}:spPr");
                spNode.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/>";
            }

            XmlNode? node = this.CreateNode($"{this._nsPrefix}:txPr");

            node.InnerXml =
                "<a:bodyPr anchorCtr=\"1\" anchor=\"ctr\" bIns=\"19050\" rIns=\"38100\" tIns=\"19050\" lIns=\"38100\" wrap=\"square\" vert=\"horz\" vertOverflow=\"ellipsis\" spcFirstLastPara=\"1\" rot=\"0\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>";
        }
    }

    ExcelTextBody _textBody;

    /// <summary>
    /// Access to text body properties
    /// </summary>
    public ExcelTextBody TextBody => this._textBody ??= new ExcelTextBody(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:txPr/a:bodyPr", this.SchemaNodeOrder);

    #endregion

    #region "Position Enum Translation"

    /// <summary>
    /// Translates the label position
    /// </summary>
    /// <param name="pos">The position enum</param>
    /// <returns>The string</returns>
    protected internal static string GetPosText(eLabelPosition pos)
    {
        switch (pos)
        {
            case eLabelPosition.Bottom:
                return "b";

            case eLabelPosition.Center:
                return "ctr";

            case eLabelPosition.InBase:
                return "inBase";

            case eLabelPosition.InEnd:
                return "inEnd";

            case eLabelPosition.Left:
                return "l";

            case eLabelPosition.Right:
                return "r";

            case eLabelPosition.Top:
                return "t";

            case eLabelPosition.OutEnd:
                return "outEnd";

            default:
                return "bestFit";
        }
    }

    /// <summary>
    /// Translates the enum position
    /// </summary>
    /// <param name="pos">The string value to translate</param>
    /// <returns>The enum value</returns>
    protected internal static eLabelPosition GetPosEnum(string pos)
    {
        switch (pos)
        {
            case "b":
                return eLabelPosition.Bottom;

            case "ctr":
                return eLabelPosition.Center;

            case "inBase":
                return eLabelPosition.InBase;

            case "inEnd":
                return eLabelPosition.InEnd;

            case "l":
                return eLabelPosition.Left;

            case "r":
                return eLabelPosition.Right;

            case "t":
                return eLabelPosition.Top;

            case "outEnd":
                return eLabelPosition.OutEnd;

            default:
                return eLabelPosition.BestFit;
        }
    }

    #endregion
}