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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// A connection point between a shape and a connection shape
/// </summary>
public class ExcelDrawingConnectionPoint : XmlHelper
{
    private readonly ExcelDrawings _drawings;
    string _path = "xdr:cxnSp/xdr:nvCxnSpPr/xdr:cNvCxnSpPr/{0}";

    internal ExcelDrawingConnectionPoint(ExcelDrawings drawings, XmlNode topNode, string elementName, string[] schemaNodeOrder)
        : base(drawings.NameSpaceManager, topNode)
    {
        this._path = string.Format(this._path, elementName);
        this._drawings = drawings;
        this.SchemaNodeOrder = schemaNodeOrder;
    }

    /// <summary>
    /// The index the connection point
    /// </summary>
    public int Index
    {
        get { return this.GetXmlNodeIntNull(this._path + "/@idx") ?? 0; }
        set
        {
            if (value <= 0)
            {
                throw new ArgumentOutOfRangeException("Index", "Index can't be negative.");
            }

            if (this._shape == null)
            {
                throw new InvalidOperationException("Can't set Index when Shape is null");
            }

            this.SetIndex(value);
        }
    }

    ExcelShape _shape = null;

    /// <summary>
    /// The shape to connect
    /// </summary>
    public ExcelShape Shape
    {
        get
        {
            if (this._shape == null)
            {
                int? id = this.GetXmlNodeIntNull(this._path + "/@id");

                if (id.HasValue)
                {
                    this._shape = this._drawings.GetById(id.Value) as ExcelShape;
                }
            }

            return this._shape;
        }
        set
        {
            if (value == null)
            {
                this.DeleteNode(this._path);
            }
            else
            {
                if (this._shape == null)
                {
                    this.SetIndex(1);
                }

                this.SetXmlNodeString(this._path + "/@id", value.Id.ToString(CultureInfo.InvariantCulture));
            }

            this._shape = value;
        }
    }

    private void SetIndex(int value)
    {
        this.SetXmlNodeString(this._path + "/@idx", value.ToString(CultureInfo.InvariantCulture));
    }
}