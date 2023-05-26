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
namespace OfficeOpenXml.Style.XmlAccess;

/// <summary>
/// Xml access class for border top level
/// </summary>
public sealed class ExcelBorderXml : StyleXmlHelper
{
    internal ExcelBorderXml(XmlNamespaceManager nameSpaceManager)
        : base(nameSpaceManager)
    {

    }
    internal ExcelBorderXml(XmlNamespaceManager nsm, XmlNode topNode) :
        base(nsm, topNode)
    {
        this._left = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(leftPath, nsm));
        this._right = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(rightPath, nsm));
        this._top = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(topPath, nsm));
        this._bottom = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(bottomPath, nsm));
        this._diagonal = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(diagonalPath, nsm));
        this._diagonalUp = this.GetBoolValue(topNode, diagonalUpPath);
        this._diagonalDown = this.GetBoolValue(topNode, diagonalDownPath);
    }
    internal override string Id
    {
        get
        {
            return this.Left.Id + this.Right.Id + this.Top.Id + this.Bottom.Id + this.Diagonal.Id + this.DiagonalUp.ToString() + this.DiagonalDown.ToString();
        }
    }
    const string leftPath = "d:left";
    ExcelBorderItemXml _left = null;
    /// <summary>
    /// Left border style properties
    /// </summary>
    public ExcelBorderItemXml Left
    {
        get
        {
            return this._left;
        }
        internal set
        {
            this._left = value;
        }
    }
    const string rightPath = "d:right";
    ExcelBorderItemXml _right = null;
    /// <summary>
    /// Right border style properties
    /// </summary>
    public ExcelBorderItemXml Right
    {
        get
        {
            return this._right;
        }
        internal set
        {
            this._right = value;
        }
    }
    const string topPath = "d:top";
    ExcelBorderItemXml _top = null;
    /// <summary>
    /// Top border style properties
    /// </summary>
    public ExcelBorderItemXml Top
    {
        get
        {
            return this._top;
        }
        internal set
        {
            this._top = value;
        }
    }
    const string bottomPath = "d:bottom";
    ExcelBorderItemXml _bottom = null;
    /// <summary>
    /// Bottom border style properties
    /// </summary>
    public ExcelBorderItemXml Bottom
    {
        get
        {
            return this._bottom;
        }
        internal set
        {
            this._bottom = value;
        }
    }
    const string diagonalPath = "d:diagonal";
    ExcelBorderItemXml _diagonal = null;
    /// <summary>
    /// Diagonal border style properties
    /// </summary>
    public ExcelBorderItemXml Diagonal
    {
        get
        {
            return this._diagonal;
        }
        internal set
        {
            this._diagonal = value;
        }
    }
    const string diagonalUpPath = "@diagonalUp";
    bool _diagonalUp = false;
    /// <summary>
    /// Diagonal up border
    /// </summary>
    public bool DiagonalUp
    {
        get
        {
            return this._diagonalUp;
        }
        internal set
        {
            this._diagonalUp = value;
        }
    }
    const string diagonalDownPath = "@diagonalDown";
    bool _diagonalDown = false;
    /// <summary>
    /// Diagonal down border
    /// </summary>
    public bool DiagonalDown
    {
        get
        {
            return this._diagonalDown;
        }
        internal set
        {
            this._diagonalDown = value;
        }
    }

    internal ExcelBorderXml Copy()
    {
        ExcelBorderXml newBorder = new ExcelBorderXml(this.NameSpaceManager);
        newBorder.Bottom = this._bottom.Copy();
        newBorder.Diagonal = this._diagonal.Copy();
        newBorder.Left = this._left.Copy();
        newBorder.Right = this._right.Copy();
        newBorder.Top = this._top.Copy();
        newBorder.DiagonalUp = this._diagonalUp;
        newBorder.DiagonalDown = this._diagonalDown;

        return newBorder;

    }

    internal override XmlNode CreateXmlNode(XmlNode topNode)
    {
        this.TopNode = topNode;
        this.CreateNode(leftPath);
        topNode.AppendChild(this._left.CreateXmlNode(this.TopNode.SelectSingleNode(leftPath, this.NameSpaceManager)));
        this.CreateNode(rightPath);
        topNode.AppendChild(this._right.CreateXmlNode(this.TopNode.SelectSingleNode(rightPath, this.NameSpaceManager)));
        this.CreateNode(topPath);
        topNode.AppendChild(this._top.CreateXmlNode(this.TopNode.SelectSingleNode(topPath, this.NameSpaceManager)));
        this.CreateNode(bottomPath);
        topNode.AppendChild(this._bottom.CreateXmlNode(this.TopNode.SelectSingleNode(bottomPath, this.NameSpaceManager)));
        this.CreateNode(diagonalPath);
        topNode.AppendChild(this._diagonal.CreateXmlNode(this.TopNode.SelectSingleNode(diagonalPath, this.NameSpaceManager)));
        if (this._diagonalUp)
        {
            this.SetXmlNodeString(diagonalUpPath, "1");
        }
        if (this._diagonalDown)
        {
            this.SetXmlNodeString(diagonalDownPath, "1");
        }
        return topNode;
    }
}