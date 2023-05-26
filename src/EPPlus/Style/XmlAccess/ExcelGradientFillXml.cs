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
using System.Globalization;
namespace OfficeOpenXml.Style.XmlAccess;

/// <summary>
/// Xml access class for gradient fillsde
/// </summary>
public sealed class ExcelGradientFillXml : ExcelFillXml
{
    internal ExcelGradientFillXml(XmlNamespaceManager nameSpaceManager)
        : base(nameSpaceManager)
    {
        this.GradientColor1 = new ExcelColorXml(nameSpaceManager);
        this.GradientColor2 = new ExcelColorXml(nameSpaceManager);
    }
    internal ExcelGradientFillXml(XmlNamespaceManager nsm, XmlNode topNode) :
        base(nsm, topNode)
    {
        this.Degree = this.GetXmlNodeDouble(_degreePath);
        this.Type = this.GetXmlNodeString(_typePath)=="path" ? ExcelFillGradientType.Path : ExcelFillGradientType.Linear;
        this.GradientColor1 = new ExcelColorXml(nsm, topNode.SelectSingleNode(_gradientColor1Path, nsm));
        this.GradientColor2 = new ExcelColorXml(nsm, topNode.SelectSingleNode(_gradientColor2Path, nsm));

        this.Top = this.GetXmlNodeDouble(_topPath);
        this.Bottom = this.GetXmlNodeDouble(_bottomPath);
        this.Left = this.GetXmlNodeDouble(_leftPath);
        this.Right = this.GetXmlNodeDouble(_rightPath);
    }
    const string _typePath = "d:gradientFill/@type";
    /// <summary>
    /// Type of gradient fill. 
    /// </summary>
    public ExcelFillGradientType Type
    {
        get;
        internal set;
    }
    const string _degreePath = "d:gradientFill/@degree";
    /// <summary>
    /// Angle of the linear gradient
    /// </summary>
    public double Degree
    {
        get;
        internal set;
    }
    const string _gradientColor1Path = "d:gradientFill/d:stop[@position=\"0\"]/d:color";
    /// <summary>
    /// Gradient color 1
    /// </summary>
    public ExcelColorXml GradientColor1 
    {
        get;
        private set;
    }
    const string _gradientColor2Path = "d:gradientFill/d:stop[@position=\"1\"]/d:color";
    /// <summary>
    /// Gradient color 2
    /// </summary>
    public ExcelColorXml GradientColor2
    {
        get;
        private set;
    }
    const string _bottomPath = "d:gradientFill/@bottom";
    /// <summary>
    /// Percentage format bottom
    /// </summary>
    public double Bottom
    { 
        get; 
        internal set; 
    }
    const string _topPath = "d:gradientFill/@top";
    /// <summary>
    /// Percentage format top
    /// </summary>
    public double Top
    {
        get;
        internal set;
    }
    const string _leftPath = "d:gradientFill/@left";
    /// <summary>
    /// Percentage format left
    /// </summary>
    public double Left
    {
        get;
        internal set;
    }
    const string _rightPath = "d:gradientFill/@right";
    /// <summary>
    /// Percentage format right
    /// </summary>
    public double Right
    {
        get;
        internal set;
    }
    internal override string Id
    {
        get
        {
            return base.Id + this.Degree.ToString() + this.GradientColor1.Id + this.GradientColor2.Id + this.Type + this.Left.ToString() + this.Right.ToString() + this.Bottom.ToString() + this.Top.ToString();
        }
    }

    #region Public Properties
    #endregion
    internal override ExcelFillXml Copy()
    {
        ExcelGradientFillXml newFill = new ExcelGradientFillXml(this.NameSpaceManager);
        newFill.PatternType = this._fillPatternType;
        newFill.BackgroundColor = this._backgroundColor.Copy();
        newFill.PatternColor = this._patternColor.Copy();

        newFill.GradientColor1 = this.GradientColor1.Copy();
        newFill.GradientColor2 = this.GradientColor2.Copy();
        newFill.Type = this.Type;
        newFill.Degree = this.Degree;
        newFill.Top = this.Top;
        newFill.Bottom = this.Bottom;
        newFill.Left = this.Left;
        newFill.Right = this.Right;
            
        return newFill;
    }

    internal override XmlNode CreateXmlNode(XmlNode topNode)
    {
        this.TopNode = topNode;
        this.CreateNode("d:gradientFill");
        if(this.Type==ExcelFillGradientType.Path)
        {
            this.SetXmlNodeString(_typePath, "path");
        }

        if(!double.IsNaN(this.Degree))
        {
            this.SetXmlNodeString(_degreePath, this.Degree.ToString(CultureInfo.InvariantCulture));
        }

        if (this.GradientColor1!=null)
        {
            /*** Gradient color node 1***/
            XmlNode? node = this.TopNode.SelectSingleNode("d:gradientFill", this.NameSpaceManager);
            XmlElement? stopNode = node.OwnerDocument.CreateElement("stop", ExcelPackage.schemaMain);
            stopNode.SetAttribute("position", "0");
            node.AppendChild(stopNode);
            XmlElement? colorNode = node.OwnerDocument.CreateElement("color", ExcelPackage.schemaMain);
            stopNode.AppendChild(colorNode);
            this.GradientColor1.CreateXmlNode(colorNode);

            /*** Gradient color node 2***/
            stopNode = node.OwnerDocument.CreateElement("stop", ExcelPackage.schemaMain);
            stopNode.SetAttribute("position", "1");
            node.AppendChild(stopNode);
            colorNode = node.OwnerDocument.CreateElement("color", ExcelPackage.schemaMain);
            stopNode.AppendChild(colorNode);

            this.GradientColor2.CreateXmlNode(colorNode);
        }
        if (!double.IsNaN(this.Top))
        {
            this.SetXmlNodeString(_topPath, this.Top.ToString("F5",CultureInfo.InvariantCulture));
        }

        if (!double.IsNaN(this.Bottom))
        {
            this.SetXmlNodeString(_bottomPath, this.Bottom.ToString("F5", CultureInfo.InvariantCulture));
        }

        if (!double.IsNaN(this.Left))
        {
            this.SetXmlNodeString(_leftPath, this.Left.ToString("F5", CultureInfo.InvariantCulture));
        }

        if (!double.IsNaN(this.Right))
        {
            this.SetXmlNodeString(_rightPath, this.Right.ToString("F5", CultureInfo.InvariantCulture));
        }

        return topNode;
    }
}