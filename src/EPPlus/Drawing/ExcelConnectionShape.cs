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
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// An Excel shape.
/// </summary>
public sealed class ExcelConnectionShape : ExcelShapeBase
{
    internal ExcelConnectionShape(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null)
        : base(drawings, node, "xdr:cxnSp", "xdr:nvCxnSpPr/xdr:cNvPr", parent) =>
        this.Init(drawings, node);

    internal ExcelConnectionShape(ExcelDrawings drawings, XmlNode node, eShapeConnectorStyle style, ExcelShape startShape, ExcelShape endShape)
        : base(drawings, node, "xdr:cxnSp", "xdr:nvCxnSpPr/xdr:cNvPr")
    {
        XmlElement shapeNode = node.OwnerDocument.CreateElement("xdr", "cxnSp", ExcelPackage.schemaSheetDrawings);
        shapeNode.SetAttribute("macro", "");
        _ = node.AppendChild(shapeNode);

        shapeNode.InnerXml = this.ShapeStartXml();
        _ = node.AppendChild(shapeNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));

        this.Init(drawings, node);
        this.ConnectionStart.Shape = startShape;
        this.ConnectionEnd.Shape = endShape;
        this.Style = style;
    }

    private void Init(ExcelDrawings drawings, XmlNode node)
    {
        this.ConnectionStart = new ExcelDrawingConnectionPoint(drawings, node, "a:stCxn", this.SchemaNodeOrder);
        this.ConnectionEnd = new ExcelDrawingConnectionPoint(drawings, node, "a:endCxn", this.SchemaNodeOrder);
    }

    #region "Public methods"

    #endregion

    #region "Private Methods"

    private string ShapeStartXml()
    {
        StringBuilder xml = new StringBuilder();

        _ = xml.AppendFormat("<xdr:nvCxnSpPr><xdr:cNvPr id=\"{0}\" name=\"{1}\" /></xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\" /></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\" /></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\" /></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\" /></a:fontRef></xdr:style>",
                         this._id,
                         this.Name);

        return xml.ToString();
    }

    #endregion

    internal override void DeleteMe()
    {
        if (this.Fill.Style == eFillStyle.BlipFill)
        {
            IPictureContainer container = this.Fill.BlipFill;
            this._drawings._package.PictureStore.RemoveImage(container.ImageHash, this.Fill.BlipFill);
        }

        base.DeleteMe();
    }

    internal new string Id => this.Name + this.Text;

    /// <summary>
    /// Connection starting point
    /// </summary>
    public ExcelDrawingConnectionPoint ConnectionStart { get; private set; }

    /// <summary>
    /// Connection ending point
    /// </summary>
    public ExcelDrawingConnectionPoint ConnectionEnd { get; private set; }

    /// <summary>
    /// Shape connector style
    /// </summary>
    public new eShapeConnectorStyle Style
    {
        get
        {
            string v = this.GetXmlNodeString(this._shapeStylePath);

            try
            {
                return (eShapeConnectorStyle)Enum.Parse(typeof(eShapeConnectorStyle), v, true);
            }
            catch
            {
                throw new Exception(string.Format("Invalid shapetype {0}", v));
            }
        }
        set => this.SetXmlNodeString(this._shapeStylePath, value.ToEnumString());
    }
}