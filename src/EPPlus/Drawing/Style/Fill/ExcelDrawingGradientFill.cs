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
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill;

/// <summary>
/// A gradient fill. This fill gradual transition from one color to the next.
/// </summary>s
public class ExcelDrawingGradientFill : ExcelDrawingFillBase
{
    private string[] _schemaNodeOrder;
    internal ExcelDrawingGradientFill(XmlNamespaceManager nsm, XmlNode topNode, string[]  schemaNodeOrder, Action initXml) : base(nsm, topNode,"", initXml)
    {
        this._schemaNodeOrder = schemaNodeOrder;
        this.GetXml();
    }

    /// <summary>
    /// The direction(s) in which to flip the gradient while tiling
    /// </summary>
    public eTileFlipMode TileFlip { get; set; }
    /// <summary>
    /// If the fill rotates along with shape.
    /// </summary>
    public bool RotateWithShape
    {
        get;
        set;
    }
    ExcelDrawingGradientFillColorList _colors = null;
    const string ColorsPath = "a:gsLst";
    /// <summary>
    /// A list of colors and their positions in percent used to generate the gradiant fill
    /// </summary>
    public ExcelDrawingGradientFillColorList Colors
    {
        get { return this._colors ??= new ExcelDrawingGradientFillColorList(this._nsm, this._topNode, ColorsPath, this._schemaNodeOrder); }
    }
    /// <summary>
    /// The fill style. 
    /// </summary>
    public override eFillStyle Style
    {
        get
        {
            return eFillStyle.GradientFill;
        }
    }

    internal override string NodeName
    {
        get
        {
            return "a:gradFill";
        }
    }

    internal override void SetXml(XmlNamespaceManager nsm, XmlNode node)
    {
        this._initXml?.Invoke();
        if (this._xml == null)
        {
            this.InitXml(nsm, node,"");
        }

        this.CheckTypeChange(this.NodeName);
        this._xml.SetXmlNodeBool("@rotWithShape", this.RotateWithShape);
        if (this.TileFlip == eTileFlipMode.None)
        {
            this._xml.DeleteNode("@flip");
        }
        else
        {
            this._xml.SetXmlNodeString("@flip", this.TileFlip.ToString().ToLower());
        }

        if (this.ShadePath==eShadePath.Linear && this.LinearSettings.Angel!=0 && this.LinearSettings.Scaled==false)
        {
            this._xml.SetXmlNodeAngel("a:lin/@ang", this.LinearSettings.Angel);
            this._xml.SetXmlNodeBool("a:lin/@scaled", this.LinearSettings.Scaled);
        }
        else if(this.ShadePath != eShadePath.Linear)
        {
            this._xml.SetXmlNodeString("a:path/@path", GetPathString(this.ShadePath));
            this._xml.SetXmlNodePercentage("a:path/a:fillToRect/@b", this.FocusPoint.BottomOffset, true, int.MaxValue/10000);
            this._xml.SetXmlNodePercentage("a:path/a:fillToRect/@t", this.FocusPoint.TopOffset, true, int.MaxValue / 10000);
            this._xml.SetXmlNodePercentage("a:path/a:fillToRect/@l", this.FocusPoint.LeftOffset, true, int.MaxValue / 10000);
            this._xml.SetXmlNodePercentage("a:path/a:fillToRect/@r", this.FocusPoint.RightOffset, true, int.MaxValue / 10000);
        }
    }

    private static string GetPathString(eShadePath shadePath)
    {
        switch(shadePath)
        {
            case eShadePath.Circle:
                return "circle";
            case eShadePath.Rectangle:
                return "rect";
            case eShadePath.Shape:
                return "shape";
            default:
                throw new ArgumentException("Unhandled ShadePath");
        }
    }

    internal override void GetXml()
    {
        this._colors = new ExcelDrawingGradientFillColorList(this._xml.NameSpaceManager, this._xml.TopNode, ColorsPath, this._schemaNodeOrder);
        this.RotateWithShape = this._xml.GetXmlNodeBool("@rotWithShape");
        try
        {
            string? s = this._xml.GetXmlNodeString("@flip");
            if (string.IsNullOrEmpty(s))
            {
                this.TileFlip = eTileFlipMode.None;
            }
            else
            {
                this.TileFlip = (eTileFlipMode)Enum.Parse(typeof(eTileFlipMode), s, true);
            }
        }
        catch
        {
            this.TileFlip = eTileFlipMode.None;
        }

        XmlNode? cols = this._xml.TopNode.SelectSingleNode("a:gsLst", this._xml.NameSpaceManager);
        if (cols != null)
        {
            foreach (XmlNode c in cols.ChildNodes)
            {
                XmlHelper? xml = XmlHelperFactory.Create(this._xml.NameSpaceManager, c);
                this._colors.Add(xml.GetXmlNodeDouble("@pos") / 1000, c);
            }
        }
        string? path= this._xml.GetXmlNodeString("a:path/@path");
        if(!string.IsNullOrEmpty(path))
        {
            if (path == "rect")
            {
                path = "rectangle";
            }

            this.ShadePath = path.ToEnum(eShadePath.Linear);
        }
        else
        {
            this.ShadePath = eShadePath.Linear;
        }
        if(this.ShadePath==eShadePath.Linear)
        {
            this.LinearSettings = new ExcelDrawingGradientFillLinearSettings(this._xml);
        }
        else
        {
            this.FocusPoint = new ExcelDrawingRectangle(this._xml, "a:path/a:fillToRect/", 0);
        }
    }
    eShadePath _shadePath = eShadePath.Linear;
    /// <summary>
    /// Specifies the shape of the path to follow
    /// </summary>
    public eShadePath ShadePath
    {
        get
        {
            return this._shadePath;
        }
        set
        {
            if(value==eShadePath.Linear)
            {
                this.LinearSettings = new ExcelDrawingGradientFillLinearSettings();
                this.FocusPoint = null;
            }
            else
            {
                this.LinearSettings = null;
                this.FocusPoint = new ExcelDrawingRectangle(50);
            }

            this._shadePath = value;
        }
    }

    /// <summary>
    /// The focuspoint when ShadePath is set to a non linear value.
    /// This property is set to null if ShadePath is set to Linear
    /// </summary>
    public ExcelDrawingRectangle FocusPoint
    {
        get;
        private set;
    }
    /// <summary>
    /// Linear gradient settings.
    /// This property is set to null if ShadePath is set to Linear
    /// </summary>
    public ExcelDrawingGradientFillLinearSettings LinearSettings
    {
        get;
        private set;
    }
    internal override void UpdateXml()
    {
        if (this._xml == null)
        {
            this.CreateXmlHelper();
        }

        this.SetXml(this._nsm, this._xml.TopNode);
            
    }
}