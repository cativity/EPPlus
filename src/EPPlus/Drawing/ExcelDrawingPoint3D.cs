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
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// A point in a 3D space
/// </summary>
public class ExcelDrawingPoint3D : XmlHelper
{
    private readonly string _xPath = "{0}/@{1}x";
    private readonly string _yPath = "{0}/@{1}y";
    private readonly string _zPath = "{0}/@{1}z";
    private readonly Action<bool> _initParent;
    bool isInit;

    internal ExcelDrawingPoint3D(XmlNamespaceManager nameSpaceManager,
                                 XmlNode topNode,
                                 string[] schemaNodeOrder,
                                 string path,
                                 string prefix,
                                 Action<bool> initParent)
        : base(nameSpaceManager, topNode)
    {
        this.SchemaNodeOrder = schemaNodeOrder;
        this._xPath = string.Format(this._xPath, path, prefix);
        this._yPath = string.Format(this._yPath, path, prefix);
        this._zPath = string.Format(this._zPath, path, prefix);
        this._initParent = initParent;
    }

    /// <summary>
    /// The X coordinate in point
    /// </summary>
    public double X
    {
        get => this.GetXmlNodeEmuToPtNull(this._xPath) ?? 0;
        set
        {
            if (this.isInit == false)
            {
                this._initParent(false);
            }

            this.SetXmlNodeEmuToPt(this._xPath, value);
        }
    }

    /// <summary>
    /// The Y coordinate
    /// </summary>
    public double Y
    {
        get => this.GetXmlNodeEmuToPtNull(this._yPath) ?? 0;
        set
        {
            if (this.isInit == false)
            {
                this._initParent(false);
            }

            this.SetXmlNodeEmuToPt(this._yPath, value);
        }
    }

    /// <summary>
    /// The Z coordinate
    /// </summary>
    public double Z
    {
        get => this.GetXmlNodeEmuToPtNull(this._zPath) ?? 0;
        set
        {
            if (this.isInit == false)
            {
                this._initParent(false);
            }

            this.SetXmlNodeEmuToPt(this._zPath, value);
        }
    }

    internal void InitXml()
    {
        if (this.isInit == false)
        {
            this.isInit = true;

            if (!this.ExistsNode(this._xPath))
            {
                this.X = 0;
                this.Y = 0;
                this.Z = 0;
            }
        }
    }
}