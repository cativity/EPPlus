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
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Effect;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
using System;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Chart.Style;

/// <summary>
/// A style entry for a chart part.
/// </summary>
public class ExcelChartStyleEntry : XmlHelper
{
    string _fillReferencePath = "{0}/cs:fillRef";
    string _borderReferencePath = "{0}/cs:lnRef ";
    string _effectReferencePath = "{0}/cs:effectRef";
    string _fontReferencePath = "{0}/cs:fontRef";
        
    string _fillPath = "{0}/cs:spPr";
    string _borderPath = "{0}/cs:spPr/a:ln";
    string _effectPath = "{0}/cs:spPr/a:effectLst";
    string _scene3DPath = "{0}/cs:spPr/a:scene3d";
    string _sp3DPath = "{0}/cs:spPr/a:sp3d";

    string _defaultTextRunPath = "{0}/cs:defRPr";
    string _defaultTextBodyPath = "{0}/cs:bodyPr";
    private readonly IPictureRelationDocument _pictureRelationDocument;
    internal ExcelChartStyleEntry(XmlNamespaceManager nsm, XmlNode topNode, string path, IPictureRelationDocument pictureRelationDocument) : base(nsm, topNode)
    {
        this.SchemaNodeOrder = new string[] { "lnRef", "fillRef", "effectRef", "fontRef", "spPr", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill","ln", "defRPr" };
        this._fillReferencePath = string.Format(this._fillReferencePath, path);
        this._borderReferencePath = string.Format(this._borderReferencePath, path);
        this._effectReferencePath = string.Format(this._effectReferencePath, path);
        this._fontReferencePath = string.Format(this._fontReferencePath, path);

        this._fillPath = string.Format(this._fillPath, path);
        this._borderPath = string.Format(this._borderPath, path);
        this._effectPath = string.Format(this._effectPath, path);
        this._scene3DPath = string.Format(this._scene3DPath, path);
        this._sp3DPath = string.Format(this._sp3DPath, path);

        this._defaultTextRunPath = string.Format(this._defaultTextRunPath, path);
        this._defaultTextBodyPath = string.Format(this._defaultTextBodyPath, path);
        this._pictureRelationDocument = pictureRelationDocument;
    }
    private ExcelChartStyleReference _borderReference = null;
    /// Border reference. 
    /// Contains an index reference to the theme and a color to be used in border styling
    public ExcelChartStyleReference BorderReference
    {
        get
        {
            return this._borderReference ??= new ExcelChartStyleReference(this.NameSpaceManager, this.TopNode, this._borderReferencePath);
        }
    }
    private ExcelChartStyleReference _fillReference = null;
    /// <summary>
    /// Fill reference. 
    /// Contains an index reference to the theme and a fill color to be used in fills
    /// </summary>
    public ExcelChartStyleReference FillReference
    {
        get
        {
            return this._fillReference ??= new ExcelChartStyleReference(this.NameSpaceManager, this.TopNode, this._fillReferencePath);
        }
    }
    private ExcelChartStyleReference _effectReference = null;
    /// <summary>
    /// Effect reference. 
    /// Contains an index reference to the theme and a color to be used in effects
    /// </summary>
    public ExcelChartStyleReference EffectReference
    {
        get
        {
            return this._effectReference ??= new ExcelChartStyleReference(this.NameSpaceManager, this.TopNode, this._effectReferencePath);
        }
    }
    ExcelChartStyleFontReference _fontReference = null;
    /// <summary>
    /// Font reference. 
    /// Contains an index reference to the theme and a color to be used for font styling
    /// </summary>
    public ExcelChartStyleFontReference FontReference
    {
        get
        {
            return this._fontReference ??= new ExcelChartStyleFontReference(this.NameSpaceManager, this.TopNode, this._fontReferencePath);
        }
    }

    private ExcelDrawingFill _fill = null;
    /// <summary>
    /// Reference to fill settings for a chart part
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            return this._fill ??= new ExcelDrawingFill(this._pictureRelationDocument,
                                                       this.NameSpaceManager,
                                                       this.TopNode,
                                                       this._fillPath,
                                                       this.SchemaNodeOrder);
        }
    }
    private ExcelDrawingBorder _border = null;
    /// <summary>
    /// Reference to border settings for a chart part
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get
        {
            return this._border ??= new ExcelDrawingBorder(this._pictureRelationDocument,
                                                           this.NameSpaceManager,
                                                           this.TopNode,
                                                           this._borderPath,
                                                           this.SchemaNodeOrder);
        }
    }
    private ExcelDrawingEffectStyle _effect = null;
    /// <summary>
    /// Reference to border settings for a chart part
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            return this._effect ??= new ExcelDrawingEffectStyle(this._pictureRelationDocument,
                                                                this.NameSpaceManager,
                                                                this.TopNode,
                                                                this._effectPath,
                                                                this.SchemaNodeOrder);
        }
    }
    private ExcelDrawing3D _threeD = null;
    /// <summary>
    /// Reference to 3D effect settings for a chart part
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get { return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, this._fillPath, this.SchemaNodeOrder); }
    }
    private ExcelTextRun _defaultTextRun = null;
    /// <summary>
    /// Reference to default text run settings for a chart part
    /// </summary>
    public ExcelTextRun DefaultTextRun
    {
        get { return this._defaultTextRun ??= new ExcelTextRun(this.NameSpaceManager, this.TopNode, this._defaultTextRunPath); }
    }
    private ExcelTextBody _defaultTextBody = null;
    /// <summary>
    /// Reference to default text body run settings for a chart part
    /// </summary>
    public ExcelTextBody DefaultTextBody
    {
        get { return this._defaultTextBody ??= new ExcelTextBody(this.NameSpaceManager, this.TopNode, this._defaultTextBodyPath); }
    }
    /// <summary>
    /// Modifier for the chart
    /// </summary>
    public eStyleEntryModifier Modifier
    {
        get
        {
            string[]? split = this.GetXmlNodeString("@mods").Split(' ');
            eStyleEntryModifier ret=0;
            foreach(string? v in split)
            {
                ret |= v.ToEnum<eStyleEntryModifier>(0);
            }
            return ret;
        }
        set
        {
            string s = "";
            foreach(eStyleEntryModifier e in Enum.GetValues(typeof(eStyleEntryModifier)))
            {
                if ((int)(value & e) != 0)
                {
                    s += e.ToString() + " ";
                }
            }
            if(s=="")
            {
                ((XmlElement)this.TopNode).RemoveAttribute("mods"); 
            }
            else
            {
                this.SetXmlNodeString("@mods", s.Substring(0,s.Length-1));
            }
        }
    }
    /// <summary>
    /// True if the entry has fill styles
    /// </summary>
    public bool HasFill
    {
        get
        {
            return this.ExistsNode($"{this._fillPath}/a:noFill") || this.ExistsNode($"{this._fillPath}/a:solidFill") || this.ExistsNode($"{this._fillPath}/a:gradFill") || this.ExistsNode($"{this._fillPath}/a:pattFill") || this.ExistsNode($"{this._fillPath}/a:blipFill");
        }
    }
    /// <summary>
    /// True if the entry has border styles
    /// </summary>
    public bool HasBorder
    {
        get
        {
            return this.ExistsNode(this._borderPath);
        }
    }
    /// <summary>
    /// True if the entry effects styles
    /// </summary>
    public bool HasEffect
    {
        get
        {
            return this.ExistsNode(this._effectPath);
        }
    }
    /// <summary>
    /// True if the entry has 3D styles
    /// </summary>
    public bool HasThreeD
    {
        get
        {
            return this.ExistsNode(this._scene3DPath) || this.ExistsNode(this._sp3DPath);
        }
    }

    /// <summary>
    /// True if the entry has text body styles
    /// </summary>
    public bool HasTextBody
    {
        get
        {
            return this.ExistsNode(this._defaultTextBodyPath);
        }
    }
    /// <summary>
    /// True if the entry has text run styles
    /// </summary>
    public bool HasTextRun
    {
        get
        {
            return this.ExistsNode(this._defaultTextRunPath);
        }
    }
}