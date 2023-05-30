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
using System.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Fill properties for drawing objects
/// </summary>
public class ExcelDrawingFill : ExcelDrawingFillBasic
{
    private readonly IPictureRelationDocument _pictureRelationDocument;

    internal ExcelDrawingFill(IPictureRelationDocument pictureRelationDocument,
                              XmlNamespaceManager nameSpaceManager,
                              XmlNode topNode,
                              string fillPath,
                              string[] schemaNodeOrder,
                              Action initXml = null)
        : base(pictureRelationDocument.Package, nameSpaceManager, topNode, fillPath, schemaNodeOrder, false, initXml)
    {
        this._pictureRelationDocument = pictureRelationDocument;

        if (this._fillNode != null)
        {
            this.LoadFill();
        }
    }

    /// <summary>
    /// Load the fill from the xml
    /// </summary>
    internal protected override void LoadFill()
    {
        this._fillTypeNode ??= this._fillNode.SelectSingleNode("a:pattFill", this.NameSpaceManager)
                               ?? this._fillNode.SelectSingleNode("a:blipFill", this.NameSpaceManager);

        if (this._fillTypeNode == null)
        {
            base.LoadFill();

            return;
        }

        switch (this._fillTypeNode.LocalName)
        {
            case "pattFill":
                this._style = eFillStyle.PatternFill;
                this._patternFill = new ExcelDrawingPatternFill(this.NameSpaceManager, this._fillTypeNode, "", this.SchemaNodeOrder, this._initXml);

                break;

            case "blipFill":
                this._style = eFillStyle.BlipFill;

                this._blipFill = new ExcelDrawingBlipFill(this._pictureRelationDocument,
                                                          this.NameSpaceManager,
                                                          this._fillTypeNode,
                                                          "",
                                                          this.SchemaNodeOrder,
                                                          this._initXml);

                break;

            default:
                base.LoadFill();

                break;
        }
    }

    internal override void SetFillProperty()
    {
        if (this._fillNode == null)
        {
            base.SetFillProperty();
        }

        this._patternFill = null;
        this._blipFill = null;

        switch (this._fillTypeNode.LocalName)
        {
            case "pattFill":
                this._patternFill = new ExcelDrawingPatternFill(this.NameSpaceManager, this._fillTypeNode, "", this.SchemaNodeOrder, this._initXml);
                this._patternFill.PatternType = eFillPatternStyle.Pct5;

                if (this._patternFill.BackgroundColor.ColorType == eDrawingColorType.None)
                {
                    this._patternFill.BackgroundColor.SetSchemeColor(eSchemeColor.Background1);
                }

                this._patternFill.ForegroundColor.SetSchemeColor(eSchemeColor.Text1);

                break;

            case "blipFill":
                this._blipFill = new ExcelDrawingBlipFill(this._pictureRelationDocument,
                                                          this.NameSpaceManager,
                                                          this._fillTypeNode,
                                                          "",
                                                          this.SchemaNodeOrder,
                                                          this._initXml);

                break;

            default:
                base.SetFillProperty();

                break;
        }
    }

    internal override void BeforeSave()
    {
        if (this._patternFill != null)
        {
            this.PatternFill.UpdateXml();
        }
        else if (this._blipFill != null)
        {
            this.BlipFill.UpdateXml();
        }
        else
        {
            base.BeforeSave();
        }
    }

    private ExcelDrawingPatternFill _patternFill;

    /// <summary>
    /// Reference pattern fill properties
    /// This property is only accessable when Type is set to PatternFill
    /// </summary>
    public ExcelDrawingPatternFill PatternFill
    {
        get { return this._patternFill; }
    }

    private ExcelDrawingBlipFill _blipFill;

    /// <summary>
    /// Reference gradient fill properties
    /// This property is only accessable when Type is set to BlipFill
    /// </summary>
    public ExcelDrawingBlipFill BlipFill
    {
        get { return this._blipFill; }
    }

    /// <summary>
    /// Disposes the object
    /// </summary>
    public new void Dispose()
    {
        base.Dispose();
        this._patternFill = null;
        this._blipFill = null;
    }
}