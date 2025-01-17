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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.ThreeD;

/// <summary>
/// Scene-level 3D properties to apply to a drawing
/// </summary>
public class ExcelDrawingScene3D : XmlHelper
{
    /// <summary>
    /// The xpath
    /// </summary>
    internal protected string _path;

    private readonly string _cameraPath = "{0}/a:camera";
    private readonly string _lightRigPath = "{0}/a:lightRig";
    private readonly string _backDropPath = "{0}/a:backdrop";

    internal ExcelDrawingScene3D(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path)
        : base(nameSpaceManager, topNode)
    {
        this._path = path;
        this.SchemaNodeOrder = schemaNodeOrder;
        this._cameraPath = string.Format(this._cameraPath, this._path);
        this._lightRigPath = string.Format(this._lightRigPath, this._path);
        this._backDropPath = string.Format(this._backDropPath, this._path);
    }

    ExcelDrawingScene3DCamera _camera;

    /// <summary>
    /// The placement and properties of the camera in the 3D scene
    /// </summary>
    public ExcelDrawingScene3DCamera Camera => this._camera ??= new ExcelDrawingScene3DCamera(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._cameraPath, this.InitXml);

    ExcelDrawingScene3DLightRig _lightRig;

    /// <summary>
    /// The light rig.
    /// When 3D is used, the light rig defines the lighting properties for the scene
    /// </summary>
    public ExcelDrawingScene3DLightRig LightRig =>
        this._lightRig ??=
            new ExcelDrawingScene3DLightRig(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._lightRigPath, this.InitXml);

    ExcelDrawingScene3DBackDrop _backDropPlane;

    /// <summary>
    /// The points and vectors contained within the backdrop define a plane in 3D space
    /// </summary>
    public ExcelDrawingScene3DBackDrop BackDropPlane =>
        this._backDropPlane ??=
            new ExcelDrawingScene3DBackDrop(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._backDropPath, this.InitXml);

    bool hasInit;

    internal void InitXml(bool delete)
    {
        if (delete)
        {
            this.DeleteNode(this._cameraPath);
            this.DeleteNode(this._lightRigPath);
            this.DeleteNode(this._backDropPath);
            this.hasInit = false;
        }
        else if (this.hasInit == false)
        {
            this.hasInit = true;

            if (!this.ExistsNode(this._cameraPath))
            {
                this.Camera.CameraType = ePresetCameraType.OrthographicFront;
                this.LightRig.RigType = eRigPresetType.ThreePt;
                this.LightRig.Direction = eLightRigDirection.Top;
            }
        }
    }
}