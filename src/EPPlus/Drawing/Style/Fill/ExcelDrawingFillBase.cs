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

namespace OfficeOpenXml.Drawing.Style.Fill;

/// <summary>
/// Base class for drawing fills
/// </summary>
public abstract class ExcelDrawingFillBase
{
    internal Action _initXml;

    /// <summary>
    /// Creates an instance of ExcelDrawingFillBase
    /// </summary>
    internal protected ExcelDrawingFillBase()
    {
    }

    /// <summary>
    /// Creates an instance of ExcelDrawingFillBase
    /// </summary>
    /// <param name="nsm">Namespace manager</param>
    /// <param name="topNode">The top node</param>
    /// <param name="fillPath">XPath to the fill</param>
    /// <param name="initXml">Xml initialize method</param>
    internal protected ExcelDrawingFillBase(XmlNamespaceManager nsm, XmlNode topNode, string fillPath, Action initXml)
    {
        this._initXml = initXml;
        this.InitXml(nsm, topNode, fillPath);
    }

    /// <summary>
    /// Type of fill
    /// </summary>
    public abstract eFillStyle Style { get; }

    /// <summary>
    /// Internal Check for type change
    /// </summary>
    /// <param name="type">The type</param>
    internal protected void CheckTypeChange(string type)
    {
        if (this._xml.TopNode.Name != type)
        {
            XmlNode? p = this._xml.TopNode.ParentNode;
            XmlElement? newNode = this._xml.TopNode.OwnerDocument.CreateElement(type, ExcelPackage.schemaDrawings);
            _ = p.ReplaceChild(newNode, this._xml.TopNode);
            this._xml.TopNode = newNode;
        }
    }

    /// <summary>
    /// The Xml helper
    /// </summary>
    internal protected XmlHelper _xml;

    /// <summary>
    /// The top node
    /// </summary>
    internal protected XmlNode _topNode;

    /// <summary>
    /// The name space manager
    /// </summary>
    internal protected XmlNamespaceManager _nsm;

    /// <summary>
    /// The XPath
    /// </summary>
    internal protected string _fillPath = "";

    /// <summary>
    /// Init xml
    /// </summary>
    /// <param name="nsm">Xml namespace manager</param>
    /// <param name="node">The node</param>
    /// <param name="fillPath">The fill path</param>
    internal protected void InitXml(XmlNamespaceManager nsm, XmlNode node, string fillPath)
    {
        this._fillPath = fillPath;
        this._nsm = nsm;
        this._topNode = node;

        if (string.IsNullOrEmpty(fillPath))
        {
            this._xml = XmlHelperFactory.Create(nsm, node);
        }
        else
        {
            this._xml = null;
        }
    }

    /// <summary>
    /// Create the Xml Helper
    /// </summary>
    protected internal void CreateXmlHelper()
    {
        this._xml = XmlHelperFactory.Create(this._nsm, this._topNode);

        this._xml.SchemaNodeOrder = new string[]
        {
            "tickLblPos", "spPr", "txPr", "dLblPos", "crossAx", "printSettings", "showVal", "prstGeom", "noFill", "solidFill", "blipFill", "gradFill",
            "noFill", "pattFill", "ln", "prstDash", "blip", "srcRect", "tile", "stretch"
        };

        this._xml.TopNode = this._xml.CreateNode(this._fillPath + "/" + this.NodeName);
    }

    internal abstract string NodeName { get; }

    internal abstract void GetXml();

    internal abstract void SetXml(XmlNamespaceManager nsm, XmlNode node);

    internal abstract void UpdateXml();
}