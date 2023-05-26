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
using System.Xml;
using OfficeOpenXml.Style;
using System.Globalization;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Utils.Extensions;
using System.Runtime.InteropServices;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// An axis for a chart
/// </summary>
public abstract class ExcelChartAxis : XmlHelper, IDrawingStyle, IStyleMandatoryProperties
{
    /// <summary>
    /// Type of axis
    /// </summary>
    internal ExcelChart _chart;
    internal string _nsPrefix;
    private readonly string _minorGridlinesPath;
    private readonly string _majorGridlinesPath;
    private readonly string _formatPath;
    private readonly string _sourceLinkedPath;

    internal ExcelChartAxis(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string nsPrefix) :
        base(nameSpaceManager, topNode)
    {
        this._chart = chart;
        this._nsPrefix = nsPrefix;
        this._formatPath = $"{this._nsPrefix}:numFmt/@formatCode";
        this._sourceLinkedPath = $"{this._nsPrefix}:numFmt/@sourceLinked";
        this._minorGridlinesPath = $"{nsPrefix}:minorGridlines";
        this._majorGridlinesPath = $"{nsPrefix}:majorGridlines";
    }
    internal abstract string Id
    {
        get;
    }
    /// <summary>
    /// Get or Sets the major tick marks for the axis. 
    /// </summary>
    public abstract eAxisTickMark MajorTickMark
    {
        get;
        set;
    }

    /// <summary>
    /// Get or Sets the minor tick marks for the axis. 
    /// </summary>
    public abstract eAxisTickMark MinorTickMark
    {
        get;
        set;
    }
    /// <summary>
    /// The type of axis
    /// </summary>
    internal abstract eAxisType AxisType
    {
        get;
    }
    /// <summary>
    /// Where the axis is located
    /// </summary>
    public abstract eAxisPosition AxisPosition
    {
        get;
        internal set;
    }
    /// <summary>
    /// Where the axis crosses
    /// </summary>
    public abstract eCrosses Crosses
    {
        get;
        set;
    }
    /// <summary>
    /// How the axis are crossed
    /// </summary>
    public abstract eCrossBetween CrossBetween
    {
        get;
        set;
    }
    /// <summary>
    /// The value where the axis cross. 
    /// Null is automatic
    /// </summary>
    public abstract double? CrossesAt
    {
        get;
        set;
    }
    /// <summary>
    /// The Numberformat used
    /// </summary>
    public string Format
    {
        get
        {
            return this.GetXmlNodeString(this._formatPath);
        }
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
        get
        {
            return this.GetXmlNodeBool(this._sourceLinkedPath);
        }
        set
        {
            this.SetXmlNodeBool(this._sourceLinkedPath, value);
        }
    }
    /// <summary>
    /// The Position of the labels
    /// </summary>
    public abstract eTickLabelPosition LabelPosition
    {
        get;
        set;
    }
    ExcelDrawingFill _fill = null;
    /// <summary>
    /// Access to fill properties
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            return this._fill ??= new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr", this.SchemaNodeOrder);
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
            return this._border ??= new ExcelDrawingBorder(this._chart,
                                                           this.NameSpaceManager,
                                                           this.TopNode,
                                                           $"{this._nsPrefix}:spPr/a:ln",
                                                           this.SchemaNodeOrder);
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
            return this._effect ??= new ExcelDrawingEffectStyle(this._chart,
                                                                this.NameSpaceManager,
                                                                this.TopNode,
                                                                $"{this._nsPrefix}:spPr/a:effectLst",
                                                                this.SchemaNodeOrder);
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
            return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr", this.SchemaNodeOrder);
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
            return this._font ??= new ExcelTextFont(this._chart,
                                                    this.NameSpaceManager,
                                                    this.TopNode,
                                                    $"{this._nsPrefix}:txPr/a:p/a:pPr/a:defRPr",
                                                    this.SchemaNodeOrder);
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
            return this._textBody ??= new ExcelTextBody(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:txPr/a:bodyPr", this.SchemaNodeOrder);
        }
    }
    void IDrawingStyleBase.CreatespPr()
    {
        this.CreatespPrNode($"{this._nsPrefix}:spPr");
    }

    /// <summary>
    /// If the axis is deleted
    /// </summary>
    public abstract bool Deleted 
    {
        get;
        set;
    }
    /// <summary>
    /// Position of the Lables
    /// </summary>
    public abstract eTickLabelPosition TickLabelPosition 
    {
        get;
        set;
    }
    /// <summary>
    /// The scaling value of the display units for the value axis
    /// </summary>
    public abstract double DisplayUnit
    {
        get;
        set;
    }
    /// <summary>
    /// Chart axis title
    /// </summary>
    internal protected ExcelChartTitle _title=null;
    /// <summary>
    /// Gives access to the charts title properties.
    /// </summary>
    public virtual ExcelChartTitle Title
    {
        get
        {                                
            return this.GetTitle();
        }
    }

    internal abstract ExcelChartTitle GetTitle();
    #region "Scaling"
    /// <summary>
    /// Minimum value for the axis.
    /// Null is automatic
    /// </summary>
    public abstract double? MinValue
    {
        get;
        set;
    }
    /// <summary>
    /// Max value for the axis.
    /// Null is automatic
    /// </summary>
    public abstract double? MaxValue
    {
        get;
        set;
    }
    /// <summary>
    /// Major unit for the axis.
    /// Null is automatic
    /// </summary>
    public abstract double? MajorUnit
    {
        get;
        set;
    }
    /// <summary>
    /// Major time unit for the axis.
    /// Null is automatic
    /// </summary>
    public abstract eTimeUnit? MajorTimeUnit
    {
        get;
        set;
    }
    /// <summary>
    /// Minor unit for the axis.
    /// Null is automatic
    /// </summary>
    public abstract double? MinorUnit
    {
        get;
        set;
    }
    /// <summary>
    /// Minor time unit for the axis.
    /// Null is automatic
    /// </summary>
    public abstract eTimeUnit? MinorTimeUnit
    {
        get;
        set;
    }
    /// <summary>
    /// The base for a logaritmic scale
    /// Null for a normal scale
    /// </summary>
    public abstract double? LogBase
    {
        get;
        set;
    }
    /// <summary>
    /// Axis orientation
    /// </summary>
    public abstract eAxisOrientation Orientation
    {
        get;
        set;
    }
    #endregion

    #region GridLines 
    ExcelDrawingBorder _majorGridlines = null; 
  
    /// <summary> 
    /// Major gridlines for the axis 
    /// </summary> 
    public ExcelDrawingBorder MajorGridlines
    { 
        get
        {
            return this._majorGridlines ??= new ExcelDrawingBorder(this._chart,
                                                                   this.NameSpaceManager,
                                                                   this.TopNode,
                                                                   $"{this._majorGridlinesPath}/{this._nsPrefix}:spPr/a:ln",
                                                                   this.SchemaNodeOrder);
        } 
    }
    ExcelDrawingEffectStyle _majorGridlineEffects = null;
    /// <summary> 
    /// Effects for major gridlines for the axis 
    /// </summary> 
    public ExcelDrawingEffectStyle MajorGridlineEffects
    {
        get
        {
            return this._majorGridlineEffects ??= new ExcelDrawingEffectStyle(this._chart,
                                                                              this.NameSpaceManager,
                                                                              this.TopNode,
                                                                              $"{this._majorGridlinesPath}/{this._nsPrefix}:spPr/a:effectLst",
                                                                              this.SchemaNodeOrder);
        }
    }

    ExcelDrawingBorder _minorGridlines = null; 
  
    /// <summary> 
    /// Minor gridlines for the axis 
    /// </summary> 
    public ExcelDrawingBorder MinorGridlines
    { 
        get
        {
            return this._minorGridlines ??= new ExcelDrawingBorder(this._chart,
                                                                   this.NameSpaceManager,
                                                                   this.TopNode,
                                                                   $"{this._minorGridlinesPath}/{this._nsPrefix}:spPr/a:ln",
                                                                   this.SchemaNodeOrder);
        } 
    }
    ExcelDrawingEffectStyle _minorGridlineEffects = null;
    /// <summary> 
    /// Effects for minor gridlines for the axis 
    /// </summary> 
    public ExcelDrawingEffectStyle MinorGridlineEffects
    {
        get
        {
            return this._minorGridlineEffects ??= new ExcelDrawingEffectStyle(this._chart,
                                                                              this.NameSpaceManager,
                                                                              this.TopNode,
                                                                              $"{this._minorGridlinesPath}/{this._nsPrefix}:spPr/a:effectLst",
                                                                              this.SchemaNodeOrder);
        }
    }
    /// <summary>
    /// True if the axis has major Gridlines
    /// </summary>
    public bool HasMajorGridlines
    {
        get
        {
            return this.ExistsNode(this._majorGridlinesPath);
        }
    }
    /// <summary>
    /// True if the axis has minor Gridlines
    /// </summary>
    public bool HasMinorGridlines
    {
        get
        {
            return this.ExistsNode(this._minorGridlinesPath);
        }
    }        
    /// <summary> 
    /// Removes Major and Minor gridlines from the Axis 
    /// </summary> 
    public void RemoveGridlines()
    {
        this.RemoveGridlines(true,true); 
    }
    /// <summary>
    ///  Removes gridlines from the Axis
    /// </summary>
    /// <param name="removeMajor">Indicates if the Major gridlines should be removed</param>
    /// <param name="removeMinor">Indicates if the Minor gridlines should be removed</param>
    public void RemoveGridlines(bool removeMajor, bool removeMinor)
    { 
        if (removeMajor) 
        {
            this.DeleteNode(this._majorGridlinesPath);
            this._majorGridlines = null; 
        } 
  
        if (removeMinor) 
        {
            this.DeleteNode(this._minorGridlinesPath);
            this._minorGridlines = null; 
        } 
    }
    /// <summary>
    /// Adds gridlines and styles them according to the style selected in the StyleManager
    /// </summary>
    /// <param name="addMajor">Indicates if the Major gridlines should be added</param>
    /// <param name="addMinor">Indicates if the Minor gridlines should be added</param>
    public void AddGridlines(bool addMajor=true, bool addMinor=false)
    {
        if(addMajor)
        {
            this.CreateNode(this._majorGridlinesPath);
            this._chart.ApplyStyleOnPart(this, this._chart._styleManager?.Style?.GridlineMajor);
        }
        if (addMinor)
        {
            this.CreateNode(this._minorGridlinesPath);
            this._chart.ApplyStyleOnPart(this, this._chart._styleManager?.Style?.GridlineMinor);
        }
    }
    /// <summary>
    /// Adds the axis title and styles it according to the style selected in the StyleManager
    /// </summary>
    /// <param name="title"></param>
    public void AddTitle(string title)
    {
        this.Title.Text = title;
        this._chart.ApplyStyleOnPart(this.Title, this._chart._styleManager?.Style?.AxisTitle);
    }
    /// <summary>
    /// Removes the axis title
    /// </summary>
    public void RemoveTitle()
    {
        this.DeleteNode($"{this._nsPrefix}:title");
    }
    /// <summary>
    /// 
    /// </summary>
    /// <param name="type"></param>
    internal void ChangeAxisType(eAxisType type)
    {
        string[]? children = CopyToSchemaNodeOrder(ExcelChartAxisStandard._schemaNodeOrderDateShared, ExcelChartAxisStandard._schemaNodeOrderDate);
        this.RenameNode(this.TopNode, "c", "dateAx", children);            
    }
    #endregion
    internal XmlNode AddTitleNode()
    {
        XmlNode? node = this.TopNode.SelectSingleNode($"{this._nsPrefix}:title", this.NameSpaceManager);
        if (node == null)
        {
            node = this.CreateNode($"{this._nsPrefix}:title");
            if (this._chart._isChartEx == false)
            {
                node.InnerXml = ExcelChartTitle.GetInitXml(this._nsPrefix);
            }
        }
        return node;
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