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

using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.ThreeD;

/// <summary>
/// Defines a bevel off a shape
/// </summary>
public class ExcelDrawing3DBevel : XmlHelper
{
    bool _isInit;
    //private string _path;
    private readonly string _widthPath = "{0}/@w";
    private readonly string _heightPath = "{0}/@h";
    private readonly string _typePath = "{0}/@prst";
    private readonly Action<bool> _initParent;

    internal ExcelDrawing3DBevel(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path, Action<bool> initParent)
        : base(nameSpaceManager, topNode)
    {
        this.SchemaNodeOrder = schemaNodeOrder;
        //this._path = path;
        this._widthPath = string.Format(this._widthPath, path);
        this._heightPath = string.Format(this._heightPath, path);
        this._typePath = string.Format(this._typePath, path);
        this._initParent = initParent;
    }

    /// <summary>
    /// The width of the bevel in points (pt)
    /// </summary>
    public double Width
    {
        get => this.GetXmlNodeEmuToPtNull(this._widthPath) ?? 6;
        set
        {
            if (!this._isInit)
            {
                this.InitXml();
            }

            this.SetXmlNodeEmuToPt(this._widthPath, value);
        }
    }

    private void InitXml()
    {
        if (this._isInit == false)
        {
            this._isInit = true;

            if (!this.ExistsNode(this._typePath))
            {
                this._initParent(false);
                this.Height = 6;
                this.Width = 6;
                this.BevelType = eBevelPresetType.Circle;
            }
        }
    }

    /// <summary>
    /// The height of the bevel in points (pt)
    /// </summary>
    public double Height
    {
        get => this.GetXmlNodeEmuToPtNull(this._heightPath) ?? 6;
        set
        {
            if (!this._isInit)
            {
                this.InitXml();
            }

            this.SetXmlNodeEmuToPt(this._heightPath, value);
        }
    }

    /// <summary>
    /// A preset bevel that can be applied to a shape.
    /// </summary>
    public eBevelPresetType BevelType
    {
        get => this.GetXmlNodeString(this._typePath).ToEnum(eBevelPresetType.Circle);
        set
        {
            if (value == eBevelPresetType.None)
            {
                this.DeleteNode(this._typePath);
                this.DeleteNode(this._heightPath);
                this.DeleteNode(this._widthPath);
            }
            else
            {
                if (!this._isInit)
                {
                    this.InitXml();
                }

                this.SetXmlNodeString(this._typePath, value.ToEnumString());
            }
        }
    }
}