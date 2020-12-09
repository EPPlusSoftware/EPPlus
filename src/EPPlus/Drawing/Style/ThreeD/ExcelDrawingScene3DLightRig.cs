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
    /// The lightrig
    /// When 3D is used, the light rig defines the lighting properties associated with the scene
    /// </summary>
    public class ExcelDrawingScene3DLightRig : XmlHelper
    {
        /// <summary>
        /// The xpath
        /// </summary>
        internal protected string _path;
        private readonly string _directionPath = "{0}/@dir";
        private readonly string _typePath = "{0}/@rig";
        private readonly string _rotationPath = "{0}/a:rot";
        private readonly Action<bool> _initParent;
        internal ExcelDrawingScene3DLightRig(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path, Action<bool> initParent) : base(nameSpaceManager, topNode)
        {
            _path = path;
            SchemaNodeOrder = schemaNodeOrder;

            _rotationPath = string.Format(_rotationPath, path);
            _directionPath = string.Format(_directionPath, path);
            _typePath = string.Format(_typePath, path);
            _initParent = initParent;
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
        /// The direction from which the light rig is oriented in relation to the scene.
        /// </summary>
        public eLightRigDirection Direction
        {
            get
            {
                return GetXmlNodeString(_directionPath).TranslateLightRigDirection();
            }
            set
            {
                _initParent(false);
                SetXmlNodeString(_directionPath, value.TranslateString());
            }
        }
        /// <summary>
        /// The preset type of light rig which is to be applied to the 3D scene
        /// </summary>
        public eRigPresetType RigType
        {
            get
            {
                return GetXmlNodeString(_typePath).ToEnum(eRigPresetType.Balanced);
            }
            set
            {
                if(value==eRigPresetType.None)
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

    }
}
