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
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.ThreeD
{
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
            if (!string.IsNullOrEmpty(path)) path += "/";
            _path = path;
            _sp3dPath = string.Format(_sp3dPath, path);
            _scene3dPath = string.Format(_scene3dPath, path);
            _bevelTPath = string.Format(_bevelTPath, _sp3dPath);
            _bevelBPath = string.Format(_bevelBPath, _sp3dPath);
            _extrusionColorPath = string.Format(_extrusionColorPath, _sp3dPath);
            _contourColorPath = string.Format(_contourColorPath, _sp3dPath);
            _extrusionHeightPath = string.Format(_extrusionHeightPath, _sp3dPath);
            _contourWidthPath = string.Format(_contourWidthPath, _sp3dPath);
            _shapeDepthPath = string.Format(_shapeDepthPath, _sp3dPath); 
            _materialTypePath = string.Format(_materialTypePath, _sp3dPath);

            AddSchemaNodeOrder(schemaNodeOrder, ExcelShapeBase._shapeNodeOrder);

            _contourColor = new ExcelDrawingColorManager(nameSpaceManager, TopNode, _contourColorPath, SchemaNodeOrder, InitContourColor);
            _extrusionColor = new ExcelDrawingColorManager(nameSpaceManager, TopNode, _extrusionColorPath, SchemaNodeOrder, InitExtrusionColor);
        }
        ExcelDrawingScene3D _scene3D = null;
        /// <summary>
        /// Defines scene-level 3D properties to apply to an object
        /// </summary>
        public ExcelDrawingScene3D Scene
        {
            get
            {
                if (_scene3D == null)
                {
                    _scene3D = new ExcelDrawingScene3D(NameSpaceManager, TopNode, SchemaNodeOrder, _scene3dPath);
                }
                return _scene3D;
            }
        }
        /// <summary>
        /// The height of the extrusion
        /// </summary>
        public double ExtrusionHeight
        {   
            get
            {
                return GetXmlNodeEmuToPtNull(_extrusionHeightPath)??0;
            }
            set
            {
                SetXmlNodeEmuToPt(_extrusionHeightPath, value);
            }
        }
        /// <summary>
        /// The height of the extrusion
        /// </summary>
        public double ContourWidth
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_contourWidthPath) ?? 0;
            }
            set
            {
                SetXmlNodeEmuToPt(_contourWidthPath, value);
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
                if(_topBevel==null)
                {
                    _topBevel = new ExcelDrawing3DBevel(NameSpaceManager, TopNode, SchemaNodeOrder, _bevelTPath, InitXml);
                }
                return _topBevel;
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
                if (_bottomBevel == null)
                {
                    _bottomBevel = new ExcelDrawing3DBevel(NameSpaceManager, TopNode, SchemaNodeOrder, _bevelBPath, InitXml);
                }
                return _bottomBevel;
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
                return _extrusionColor;                
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
                return _contourColor;
            }
        }
        /// <summary>
        /// The surface appearance of a shape
        /// </summary>
        public ePresetMaterialType MaterialType
        {
            get
            {
                return GetXmlNodeString(_materialTypePath).ToEnum(ePresetMaterialType.WarmMatte);
            }
            set
            {
                InitXml(false);
                SetXmlNodeString(_materialTypePath, value.ToEnumString());
            }
        }
        /// <summary>
        /// The z coordinate for the 3D shape
        /// </summary>
        public double? ShapeDepthZCoordinate
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_shapeDepthPath) ?? 0;
            }
            set
            {
                SetXmlNodeEmuToPt(_shapeDepthPath, value);
            }
        }

        internal XmlElement Scene3DElement
        {
            get
            {
                return GetNode(_scene3dPath) as XmlElement;
            }
        }
        internal XmlElement Sp3DElement
        {
            get
            {
                return GetNode(_sp3dPath) as XmlElement;
            }
        }
        bool isInit = false;
        internal void InitXml(bool delete)
        {
            if(delete)
            {
                Delete();
            }
            else
            {
                if (isInit == false)
                {
                    if (!ExistNode(_sp3dPath))
                    {
                        CreateNode(_sp3dPath);
                        Scene.InitXml(false);
                    }
                }
            }
        }
        /// <summary>
        /// Remove all 3D settings
        /// </summary>
        public void Delete()
        {
            DeleteNode(_scene3dPath);
            DeleteNode(_sp3dPath);
        }
        private void InitContourColor()
        {
            if (ContourWidth <= 0) ContourWidth = 1;
        }
        private void InitExtrusionColor()
        {
            if (ExtrusionHeight <= 0) ExtrusionHeight = 1;
        }

        internal void SetFromXml(XmlElement copyFromScene3DElement, XmlElement copyFromSp3DElement)
        {
            if(copyFromScene3DElement!=null)
            {
                var scene3DElement = (XmlElement)CreateNode(_scene3dPath);
                CopyXml(copyFromScene3DElement, scene3DElement);
            }
            if (copyFromSp3DElement!=null)
            {
                var sp3DElement = (XmlElement)CreateNode(_sp3dPath);
                CopyXml(copyFromSp3DElement, sp3DElement);
            }
        }

        private void CopyXml(XmlElement copyFrom, XmlElement to)
        {
            foreach (XmlAttribute a in copyFrom.Attributes)
            {
                to.SetAttribute(a.Name, a.NamespaceURI, a.Value);
            }
            to.InnerXml = copyFrom.InnerXml;
        }
    }
}
