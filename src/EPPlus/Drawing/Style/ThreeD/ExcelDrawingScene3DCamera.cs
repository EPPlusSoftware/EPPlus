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
    /// Settings for the camera in the 3D scene
    /// </summary>
    public class ExcelDrawingScene3DCamera : XmlHelper
    {
        /// <summary>
        /// The XPath
        /// </summary>
        internal protected string _path;
        private readonly string _fieldOfViewAnglePath = "{0}/@pov";
        private readonly string _typePath = "{0}/@prst";
        private readonly string _zoomPath = "{0}/@zoom";
        private readonly string _rotationPath = "{0}/a:rot";

        private readonly Action<bool> _initParent;
        internal ExcelDrawingScene3DCamera(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path, Action<bool> initParent) : base(nameSpaceManager, topNode)
        {
            _path = path;
            SchemaNodeOrder = schemaNodeOrder;
            _initParent = initParent;
            _rotationPath = string.Format(_rotationPath, path);
            _fieldOfViewAnglePath = string.Format(_fieldOfViewAnglePath, path);
            _typePath = string.Format(_typePath, path);
            _zoomPath = string.Format(_zoomPath, path);
        }
        ExcelDrawingSphereCoordinate _rotation = null;
        /// <summary>
        /// Defines a rotation in 3D space
        /// </summary>
        public ExcelDrawingSphereCoordinate Rotation
        {
            get
            {
                if(_rotation==null)
                {
                    _rotation = new ExcelDrawingSphereCoordinate(NameSpaceManager, TopNode, _rotationPath, _initParent);
                }
                return _rotation;
            }
        }
        /// <summary>
        /// An override for the default field of view for the camera.
        /// </summary>
        public double FieldOfViewAngle
        {
            get
            {
                return GetXmlNodeAngle(_fieldOfViewAnglePath, 0);
            }
            set
            {
                _initParent(false);
                SetXmlNodeAngle(_fieldOfViewAnglePath, value, "FieldOfViewAngle", 0, 180);
            }
        }
        /// <summary>
        /// The preset camera type that is being used.
        /// </summary>
        public ePresetCameraType CameraType
        {
            get
            {
                return GetXmlNodeString(_typePath).ToEnum(ePresetCameraType.None);
            }
            set
            {
                if(value==ePresetCameraType.None)
                {
                    _initParent(true);
                }
                else
                {
                    _initParent(false);
                    SetXmlNodeString(_typePath, value.ToEnumString());
                }
            }
        }
        /// <summary>
        /// The zoom factor of a given camera
        /// </summary>
        public double Zoom
        {
            get
            {
                return GetXmlNodePercentage(_zoomPath) ?? 100;
            }
            set
            {
                SetXmlNodePercentage(_zoomPath, value, false);
                _initParent(false);
            }
        }
    }
}
