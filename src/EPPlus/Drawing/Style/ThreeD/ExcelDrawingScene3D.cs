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

namespace OfficeOpenXml.Drawing.Style.ThreeD
{
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

        internal ExcelDrawingScene3D(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode)
        {
            _path = path;
            SchemaNodeOrder = schemaNodeOrder;
            _cameraPath = string.Format(_cameraPath, _path);
            _lightRigPath = string.Format(_lightRigPath, _path);
            _backDropPath = string.Format(_backDropPath, _path);
        }
        ExcelDrawingScene3DCamera _camera = null;
        /// <summary>
        /// The placement and properties of the camera in the 3D scene
        /// </summary>
        public ExcelDrawingScene3DCamera Camera
        {
            get
            {
                if(_camera == null)
                {
                    _camera = new ExcelDrawingScene3DCamera(NameSpaceManager, TopNode, SchemaNodeOrder, _cameraPath, InitXml);
                }
                return _camera;
            }
        }
        ExcelDrawingScene3DLightRig _lightRig = null;
        /// <summary>
        /// The light rig.
        /// When 3D is used, the light rig defines the lighting properties for the scene
        /// </summary>
        public ExcelDrawingScene3DLightRig LightRig
        {
            get
            {
                if (_lightRig == null)
                {
                    _lightRig = new ExcelDrawingScene3DLightRig(NameSpaceManager, TopNode, SchemaNodeOrder, _lightRigPath, InitXml);
                }
                return _lightRig;
            }
        }
        ExcelDrawingScene3DBackDrop _backDropPlane = null;
        /// <summary>
        /// The points and vectors contained within the backdrop define a plane in 3D space
        /// </summary>
        public ExcelDrawingScene3DBackDrop BackDropPlane
        {
            get
            {
                if (_backDropPlane == null)
                {
                    _backDropPlane = new ExcelDrawingScene3DBackDrop(NameSpaceManager, TopNode, SchemaNodeOrder, _backDropPath, InitXml);
                }
                return _backDropPlane;
            }
        }
        bool hasInit = false;
        internal void InitXml(bool delete)
        {
            
            if(delete)
            {
                DeleteNode(_cameraPath);
                DeleteNode(_lightRigPath);
                DeleteNode(_backDropPath);
                hasInit = false;
            }
            else if(hasInit==false)
            {
                hasInit = true;
                if (!ExistNode(_cameraPath))
                {
                    Camera.CameraType = ePresetCameraType.OrthographicFront;
                    LightRig.RigType = eRigPresetType.ThreePt;
                    LightRig.Direction = eLightRigDirection.Top;
                }
            }
        }
    }
}
