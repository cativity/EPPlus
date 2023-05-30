using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Base class for vml form controls
/// </summary>
public class ExcelVmlDrawingControl : ExcelVmlDrawingBase
{
    ExcelWorksheet _ws;

    internal ExcelVmlDrawingControl(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns)
        : base(topNode, ns)
    {
        this._ws = ws;
    }

    /// <summary>
    /// The Text
    /// </summary>
    public string Text
    {
        get { return this.GetXmlNodeString("v:textbox/d:div/d:font"); }
        set { this.SetXmlNodeString("v:textbox/div/font", value); }
    }

    /// <summary>
    /// Item height for an individual item
    /// </summary>
    internal int? Dx
    {
        get { return this.GetXmlNodeIntNull("x:ClientData/x:Dx"); }
        set { this.SetXmlNodeInt("x:ClientData/x:Dx", value); }
    }

    /// <summary>
    /// Number of items in a listbox.
    /// </summary>
    internal int? Page
    {
        get { return this.GetXmlNodeIntNull("x:ClientData/x:Page"); }
        set { this.SetXmlNodeInt("x:ClientData/x:Page", value); }
    }

    internal ExcelVmlDrawingFill _fill;

    internal ExcelVmlDrawingFill GetFill()
    {
        return this._fill ??= new ExcelVmlDrawingFill(this._ws.Drawings, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
    }
}