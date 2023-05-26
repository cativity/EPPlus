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
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.ThreeD;

/// <summary>
/// 3D settings for a drawing object
/// </summary>
public class ExcelDrawing3D : XmlHelper
{
    private readonly string _sp3dPath = "{0}a:sp3d";
    private readonly string _scene3dPath = "{0}a:scene3d";
    private readonly string _bevelTPath = "{0}/a:bevelT";
    private readonly string _bevelBPath = "{0}/a:bevelB";
    private readonly string _extrusionColorPath = "{0}/a:extrusionClr";
    private readonly string _contourColorPath = "{0}/a:contourClr";        
    private readonly string _contourWidthPath = "{0}/@contourW";
    private readonly string _extrusionHeightPath = "{0}/@extrusionH";
    private readonly string _shapeDepthPath = "{0}/@z";
    private readonly string _materialTypePath = "{0}/@prstMaterial";
    private readonly string _path;
    internal ExcelDrawing3D(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder) : base(nameSpaceManager, topNode)
    {
        if (!string.IsNullOrEmpty(path))
        {
            path += "/";
        }

        this._path = path;
        this._sp3dPath = string.Format(this._sp3dPath, path);
        this._scene3dPath = string.Format(this._scene3dPath, path);
        this._bevelTPath = string.Format(this._bevelTPath, this._sp3dPath);
        this._bevelBPath = string.Format(this._bevelBPath, this._sp3dPath);
        this._extrusionColorPath = string.Format(this._extrusionColorPath, this._sp3dPath);
        this._contourColorPath = string.Format(this._contourColorPath, this._sp3dPath);
        this._extrusionHeightPath = string.Format(this._extrusionHeightPath, this._sp3dPath);
        this._contourWidthPath = string.Format(this._contourWidthPath, this._sp3dPath);
        this._shapeDepthPath = string.Format(this._shapeDepthPath, this._sp3dPath);
        this._materialTypePath = string.Format(this._materialTypePath, this._sp3dPath);

        this.AddSchemaNodeOrder(schemaNodeOrder, ExcelShapeBase._shapeNodeOrder);

        this._contourColor = new ExcelDrawingColorManager(nameSpaceManager, this.TopNode, this._contourColorPath, this.SchemaNodeOrder, this.InitContourColor);
        this._extrusionColor = new ExcelDrawingColorManager(nameSpaceManager, this.TopNode, this._extrusionColorPath, this.SchemaNodeOrder, this.InitExtrusionColor);
    }
    ExcelDrawingScene3D _scene3D = null;
    /// <summary>
    /// Defines scene-level 3D properties to apply to an object
    /// </summary>
    public ExcelDrawingScene3D Scene
    {
        get
        {
            return this._scene3D ??= new ExcelDrawingScene3D(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._scene3dPath);
        }
    }
    /// <summary>
    /// The height of the extrusion
    /// </summary>
    public double ExtrusionHeight
    {   
        get
        {
            return this.GetXmlNodeEmuToPtNull(this._extrusionHeightPath)??0;
        }
        set
        {
            this.SetXmlNodeEmuToPt(this._extrusionHeightPath, value);
        }
    }
    /// <summary>
    /// The height of the extrusion
    /// </summary>
    public double ContourWidth
    {
        get
        {
            return this.GetXmlNodeEmuToPtNull(this._contourWidthPath) ?? 0;
        }
        set
        {
            this.SetXmlNodeEmuToPt(this._contourWidthPath, value);
        }
    }
    ExcelDrawing3DBevel _topBevel = null;
    /// <summary>
    /// The bevel on the top or front face of a shape
    /// </summary>
    public ExcelDrawing3DBevel TopBevel
    {
        get
        {
            return this._topBevel ??= new ExcelDrawing3DBevel(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._bevelTPath, this.InitXml);
        }
    }
    ExcelDrawing3DBevel _bottomBevel = null;
    /// <summary>
    /// The bevel on the top or front face of a shape
    /// </summary>
    public ExcelDrawing3DBevel BottomBevel
    {
        get
        {
            return this._bottomBevel ??= new ExcelDrawing3DBevel(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._bevelBPath, this.InitXml);
        }
    }
    ExcelDrawingColorManager _extrusionColor = null;
    /// <summary>
    /// The color of the extrusion applied to a shape
    /// </summary>
    public ExcelDrawingColorManager ExtrusionColor
    {
        get
        {
            return this._extrusionColor;                
        }
    }

    ExcelDrawingColorManager _contourColor = null;
    /// <summary>
    /// The color for the contour on a shape
    /// </summary>
    public ExcelDrawingColorManager ContourColor
    {
        get
        {
            return this._contourColor;
        }
    }
    /// <summary>
    /// The surface appearance of a shape
    /// </summary>
    public ePresetMaterialType MaterialType
    {
        get
        {
            return this.GetXmlNodeString(this._materialTypePath).ToEnum(ePresetMaterialType.WarmMatte);
        }
        set
        {
            this.InitXml(false);
            this.SetXmlNodeString(this._materialTypePath, value.ToEnumString());
        }
    }
    /// <summary>
    /// The z coordinate for the 3D shape
    /// </summary>
    public double? ShapeDepthZCoordinate
    {
        get
        {
            return this.GetXmlNodeEmuToPtNull(this._shapeDepthPath) ?? 0;
        }
        set
        {
            this.SetXmlNodeEmuToPt(this._shapeDepthPath, value);
        }
    }

    internal XmlElement Scene3DElement
    {
        get
        {
            return this.GetNode(this._scene3dPath) as XmlElement;
        }
    }
    internal XmlElement Sp3DElement
    {
        get
        {
            return this.GetNode(this._sp3dPath) as XmlElement;
        }
    }
    bool isInit = false;
    internal void InitXml(bool delete)
    {
        if(delete)
        {
            this.Delete();
        }
        else
        {
            if (this.isInit == false)
            {
                if (!this.ExistsNode(this._sp3dPath))
                {
                    this.CreateNode(this._sp3dPath);
                    this.Scene.InitXml(false);
                }
            }
        }
    }
    /// <summary>
    /// Remove all 3D settings
    /// </summary>
    public void Delete()
    {
        this.DeleteNode(this._scene3dPath);
        this.DeleteNode(this._sp3dPath);
    }
    private void InitContourColor()
    {
        if (this.ContourWidth <= 0)
        {
            this.ContourWidth = 1;
        }
    }
    private void InitExtrusionColor()
    {
        if (this.ExtrusionHeight <= 0)
        {
            this.ExtrusionHeight = 1;
        }
    }

    internal void SetFromXml(XmlElement copyFromScene3DElement, XmlElement copyFromSp3DElement)
    {
        if(copyFromScene3DElement!=null)
        {
            XmlElement? scene3DElement = (XmlElement)this.CreateNode(this._scene3dPath);
            CopyXml(copyFromScene3DElement, scene3DElement);
        }
        if (copyFromSp3DElement!=null)
        {
            XmlElement? sp3DElement = (XmlElement)this.CreateNode(this._sp3dPath);
            CopyXml(copyFromSp3DElement, sp3DElement);
        }
    }

    private static void CopyXml(XmlElement copyFrom, XmlElement to)
    {
        foreach (XmlAttribute a in copyFrom.Attributes)
        {
            to.SetAttribute(a.Name, a.NamespaceURI, a.Value);
        }
        to.InnerXml = copyFrom.InnerXml;
    }
}