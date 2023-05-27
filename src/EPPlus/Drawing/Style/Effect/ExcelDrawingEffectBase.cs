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

namespace OfficeOpenXml.Drawing.Style.Effect;

/// <summary>
/// Base class for all drawing effects
/// </summary>
public abstract class ExcelDrawingEffectBase : XmlHelper
{
    internal string _path;

    internal ExcelDrawingEffectBase(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path)
        : base(nameSpaceManager, topNode)
    {
        this._path = path;
        this.SchemaNodeOrder = schemaNodeOrder;
    }

    /// <summary>
    /// Completely remove the xml node, resetting the properties to it's default values.
    /// </summary>
    public void Delete()
    {
        XmlNode? node = this.GetNode(this._path);

        if (node != null)
        {
            this.TopNode.RemoveChild(node);
        }
    }
}