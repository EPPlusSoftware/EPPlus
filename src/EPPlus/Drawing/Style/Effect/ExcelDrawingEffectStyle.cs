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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect
{
    /// <summary>
    /// Effect styles of a drawing object
    /// </summary>
    public  class ExcelDrawingEffectStyle : XmlHelper
    {
        private readonly string _path;
        private readonly string _softEdgeRadiusPath = "{0}a:softEdge/@rad";
        private readonly string _blurPath = "{0}a:blur";
        private readonly string _fillOverlayPath = "{0}a:fillOverlay";
        private readonly string _glowPath = "{0}a:glow";
        private readonly string _innerShadowPath = "{0}a:innerShdw";
        private readonly string _outerShadowPath = "{0}a:outerShdw";
        private readonly string _presetShadowPath = "{0}a:prstShdw";
        private readonly string _reflectionPath = "{0}a:reflection";
        private readonly IPictureRelationDocument _pictureRelationDocument;
        internal ExcelDrawingEffectStyle(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder) : base(nameSpaceManager, topNode)
        {
            _path = path;
            if (path.Length > 0 && !path.EndsWith("/")) path += "/";
            _softEdgeRadiusPath = string.Format(_softEdgeRadiusPath, path);
            _blurPath = string.Format(_blurPath, path);
            _fillOverlayPath = string.Format(_fillOverlayPath, path);
            _glowPath = string.Format(_glowPath, path);
            _innerShadowPath = string.Format(_innerShadowPath, path);
            _outerShadowPath = string.Format(_outerShadowPath, path);
            _presetShadowPath = string.Format(_presetShadowPath, path);
            _reflectionPath = string.Format(_reflectionPath, path);
            _pictureRelationDocument = pictureRelationDocument;

            AddSchemaNodeOrder(schemaNodeOrder, ExcelShapeBase._shapeNodeOrder);   
        }
        ExcelDrawingBlurEffect _blur = null;
        /// <summary>
        /// The blur effect
        /// </summary>
        public ExcelDrawingBlurEffect Blur
        {
            get
            {
                if (_blur == null)
                {
                    _blur = new ExcelDrawingBlurEffect(NameSpaceManager, TopNode, SchemaNodeOrder, _blurPath);
                }
                return _blur;
            }
        }

        ExcelDrawingFillOverlayEffect _fillOverlay = null;
        /// <summary>
        /// The fill overlay effect. A fill overlay can be used to specify an additional fill for a drawing and blend the two together.
        /// </summary>
        public ExcelDrawingFillOverlayEffect FillOverlay
        {
            get
            {
                if(_fillOverlay==null)
                {
                    _fillOverlay = new ExcelDrawingFillOverlayEffect(_pictureRelationDocument, NameSpaceManager, TopNode, SchemaNodeOrder, _fillOverlayPath);
                }
                return _fillOverlay;
            }
        }
        ExcelDrawingGlowEffect _glow = null;
        /// <summary>
        /// The glow effect. A color blurred outline is added outside the edges of the drawing
        /// </summary>
        public ExcelDrawingGlowEffect Glow
        {
            get
            {
                if (_glow == null)
                {
                    _glow = new ExcelDrawingGlowEffect(NameSpaceManager, TopNode, SchemaNodeOrder, _glowPath);
                }
                return _glow;
            }
        }
        ExcelDrawingInnerShadowEffect _innerShadowEffect = null;
        /// <summary>
        /// The inner shadow effect. A shadow is applied within the edges of the drawing.
        /// </summary>
        public ExcelDrawingInnerShadowEffect InnerShadow
        {
            get
            {
                if (_innerShadowEffect == null)
                {
                    _innerShadowEffect = new ExcelDrawingInnerShadowEffect(NameSpaceManager, TopNode, SchemaNodeOrder, _innerShadowPath);
                }
                return _innerShadowEffect;
            }
        }
        ExcelDrawingOuterShadowEffect _outerShadow=null;
        /// <summary>
        /// The outer shadow effect. A shadow is applied outside the edges of the drawing.
        /// </summary>
        public ExcelDrawingOuterShadowEffect OuterShadow
        {
            get
            {
                if (_outerShadow == null)
                {
                    _outerShadow = new ExcelDrawingOuterShadowEffect(NameSpaceManager, TopNode, SchemaNodeOrder, _outerShadowPath);
                }
                return _outerShadow;
            }
        }
        ExcelDrawingPresetShadowEffect _presetShadow;
        /// <summary>
        /// The preset shadow effect.
        /// </summary>
        public ExcelDrawingPresetShadowEffect PresetShadow
        {
            get
            {
                if (_presetShadow == null)
                {
                    _presetShadow = new ExcelDrawingPresetShadowEffect(NameSpaceManager, TopNode, SchemaNodeOrder, _presetShadowPath);
                }
                return _presetShadow;
            }
        }
        ExcelDrawingReflectionEffect _reflection;
        /// <summary>
        /// The reflection effect.
        /// </summary>
        public ExcelDrawingReflectionEffect Reflection
        {
            get
            {
                if (_reflection == null)
                {
                    _reflection = new ExcelDrawingReflectionEffect(NameSpaceManager, TopNode, SchemaNodeOrder, _reflectionPath);
                }
                return _reflection;
            }
        }
        /// <summary>
        /// Soft edge radius. A null value indicates no radius
        /// </summary>
        public double? SoftEdgeRadius
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_softEdgeRadiusPath);
            }
            set
            {
                if (value.HasValue)
                {
                    SetXmlNodeEmuToPt(_softEdgeRadiusPath, value.Value);
                }
                else
                {
                    DeleteNode(_softEdgeRadiusPath, true);
                }
            }
        }        
        internal XmlElement EffectElement
        {
            get
            {
                if(string.IsNullOrEmpty(_path))
                {
                    return (XmlElement)TopNode;
                }
                if (ExistNode(_path))
                {
                    return (XmlElement)GetNode(_path);
                }
                return (XmlElement)CreateNode(_path);
            }
        }
        /// <summary>
        /// If the drawing has any inner shadow properties set
        /// </summary>
        public bool HasInnerShadow
        {
            get
            {
                return ExistNode(_innerShadowPath);
            }
        }
        /// <summary>
        /// If the drawing has any outer shadow properties set
        /// </summary>
        public bool HasOuterShadow
        {
            get
            {
                return ExistNode(_outerShadowPath);
            }
        }
        /// <summary>
        /// If the drawing has any preset shadow properties set
        /// </summary>
        public bool HasPresetShadow
        {
            get
            {
                return ExistNode(_presetShadowPath);
            }
        }
        /// <summary>
        /// If the drawing has any blur properties set
        /// </summary>
        public bool HasBlur
        {
            get
            {
                return ExistNode(_blurPath);
            }
        }
        /// <summary>
        /// If the drawing has any glow properties set
        /// </summary>
        public bool HasGlow
        {
            get
            {
                return ExistNode(_glowPath);
            }
        }
        /// <summary>
        /// If the drawing has any fill overlay properties set
        /// </summary>
        public bool HasFillOverlay
        {
            get
            {
                return ExistNode(_fillOverlayPath);
            }
        }

        internal void SetFromXml(XmlElement copyFromEffectElement)
        {
            XmlElement effectElement = EffectElement;

            foreach (XmlAttribute a in copyFromEffectElement.Attributes)
            {
                effectElement.SetAttribute(a.Name, a.NamespaceURI, a.Value);
            }
            effectElement.InnerXml = copyFromEffectElement.InnerXml;
        }
        #region Private Methods
        private void SetPredefinedOuterShadow(ePresetExcelShadowType shadowType)
        {
            OuterShadow.Color.SetPresetColor(ePresetColor.Black);

            switch (shadowType)
            {
                case ePresetExcelShadowType.PerspectiveUpperLeft:
                    OuterShadow.Color.Transforms.AddAlpha(20);
                    OuterShadow.BlurRadius = 6;
                    OuterShadow.Distance = 0;
                    OuterShadow.Direction = 225;
                    OuterShadow.Alignment = eRectangleAlignment.BottomRight;
                    OuterShadow.HorizontalSkewAngle = 20;
                    OuterShadow.VerticalScalingFactor = 23;
                    break;
                case ePresetExcelShadowType.PerspectiveUpperRight:
                    OuterShadow.Color.Transforms.AddAlpha(20);
                    OuterShadow.BlurRadius = 6;
                    OuterShadow.Distance = 0;
                    OuterShadow.Direction = 315;
                    OuterShadow.Alignment = eRectangleAlignment.BottomLeft;
                    OuterShadow.HorizontalSkewAngle = -20;
                    OuterShadow.VerticalScalingFactor = 23;
                    break;
                case ePresetExcelShadowType.PerspectiveBelow:
                    OuterShadow.Color.Transforms.AddAlpha(15);
                    OuterShadow.BlurRadius = 12;
                    OuterShadow.Distance = 25;
                    OuterShadow.Direction = 90;
                    OuterShadow.HorizontalScalingFactor = 90;
                    OuterShadow.VerticalScalingFactor = -19;
                    break;
                case ePresetExcelShadowType.PerspectiveLowerLeft:
                    OuterShadow.Color.Transforms.AddAlpha(20);
                    OuterShadow.BlurRadius = 6;
                    OuterShadow.Distance = 1;
                    OuterShadow.Direction = 135;
                    OuterShadow.Alignment = eRectangleAlignment.BottomRight;
                    OuterShadow.HorizontalSkewAngle = 13.34;
                    OuterShadow.VerticalScalingFactor = -23;
                    break;
                case ePresetExcelShadowType.PerspectiveLowerRight:
                    OuterShadow.Color.Transforms.AddAlpha(20);
                    OuterShadow.BlurRadius = 6;
                    OuterShadow.Distance = 1;
                    OuterShadow.Direction = 45;
                    OuterShadow.Alignment = eRectangleAlignment.BottomLeft;
                    OuterShadow.HorizontalSkewAngle = -13.34;
                    OuterShadow.VerticalScalingFactor = -23;
                    break;

                case ePresetExcelShadowType.OuterCenter:
                    OuterShadow.Color.Transforms.AddAlpha(40);
                    OuterShadow.VerticalScalingFactor = 102;
                    OuterShadow.HorizontalScalingFactor = 102;
                    OuterShadow.BlurRadius = 5;
                    OuterShadow.Alignment = eRectangleAlignment.Center;
                    break;
                default:
                    OuterShadow.Color.Transforms.AddAlpha(40);
                    OuterShadow.BlurRadius = 4;
                    OuterShadow.Distance = 3;
                    switch (shadowType)
                    {
                        case ePresetExcelShadowType.OuterTopLeft:
                            OuterShadow.Direction = 225;
                            OuterShadow.Alignment = eRectangleAlignment.BottomRight;
                            break;
                        case ePresetExcelShadowType.OuterTop:
                            OuterShadow.Direction = 270;
                            OuterShadow.Alignment = eRectangleAlignment.Bottom;
                            break;
                        case ePresetExcelShadowType.OuterTopRight:
                            OuterShadow.Direction = 315;
                            OuterShadow.Alignment = eRectangleAlignment.BottomLeft;
                            break;
                        case ePresetExcelShadowType.OuterLeft:
                            OuterShadow.Direction = 180;
                            OuterShadow.Alignment = eRectangleAlignment.Right;
                            break;
                        case ePresetExcelShadowType.OuterRight:
                            OuterShadow.Direction = 0;
                            OuterShadow.Alignment = eRectangleAlignment.Left;
                            break;
                        case ePresetExcelShadowType.OuterBottomLeft:
                            OuterShadow.Direction = 135;
                            OuterShadow.Alignment = eRectangleAlignment.TopRight;
                            break;
                        case ePresetExcelShadowType.OuterBottom:
                            OuterShadow.Direction = 90;
                            OuterShadow.Alignment = eRectangleAlignment.Top;
                            break;
                        case ePresetExcelShadowType.OuterBottomRight:
                            OuterShadow.Direction = 45;
                            OuterShadow.Alignment = eRectangleAlignment.TopLeft;
                            break;
                    }
                    break;
            }

            OuterShadow.RotateWithShape = false;

        }
        private void SetPredefinedInnerShadow(ePresetExcelShadowType shadowType)
        {
            InnerShadow.Color.SetPresetColor(ePresetColor.Black);

            if (shadowType == ePresetExcelShadowType.InnerCenter)
            {
                InnerShadow.Color.Transforms.AddAlpha(0);
                InnerShadow.Direction = 0;
                InnerShadow.Distance = 0;
                InnerShadow.BlurRadius = 9;
            }
            else
            {
                InnerShadow.Color.Transforms.AddAlpha(50);
                InnerShadow.BlurRadius = 5;
                InnerShadow.Distance = 4;
            }

            switch (shadowType)
            {
                case ePresetExcelShadowType.InnerTopLeft:
                    InnerShadow.Direction = 225;
                    break;
                case ePresetExcelShadowType.InnerTop:
                    InnerShadow.Direction = 270;
                    break;
                case ePresetExcelShadowType.InnerTopRight:
                    InnerShadow.Direction = 315;
                    break;
                case ePresetExcelShadowType.InnerLeft:
                    InnerShadow.Direction = 180;
                    break;
                case ePresetExcelShadowType.InnerRight:
                    InnerShadow.Direction = 0;
                    break;
                case ePresetExcelShadowType.InnerBottomLeft:
                    InnerShadow.Direction = 135;
                    break;
                case ePresetExcelShadowType.InnerBottom:
                    InnerShadow.Direction = 90;
                    break;
                case ePresetExcelShadowType.InnerBottomRight:
                    InnerShadow.Direction = 45;
                    break;
            }
        }
        /// <summary>
        /// Set a predefined glow matching the preset types in Excel
        /// </summary>
        /// <param name="softEdgesType">The preset type</param>
        public void SetPresetSoftEdges(ePresetExcelSoftEdgesType softEdgesType)
        {
            switch (softEdgesType)
            {
                case ePresetExcelSoftEdgesType.None:
                    SoftEdgeRadius = null;
                    break;
                case ePresetExcelSoftEdgesType.SoftEdge1Pt:
                    SoftEdgeRadius = 1;
                    break;
                case ePresetExcelSoftEdgesType.SoftEdge2_5Pt:
                    SoftEdgeRadius = 2.5;
                    break;
                case ePresetExcelSoftEdgesType.SoftEdge5Pt:
                    SoftEdgeRadius = 5;
                    break;
                case ePresetExcelSoftEdgesType.SoftEdge10Pt:
                    SoftEdgeRadius = 10;
                    break;
                case ePresetExcelSoftEdgesType.SoftEdge25Pt:
                    SoftEdgeRadius = 25;
                    break;
                case ePresetExcelSoftEdgesType.SoftEdge50Pt:
                    SoftEdgeRadius = 50;
                    break;
            }
        }


        /// <summary>
        /// Set a predefined glow matching the preset types in Excel
        /// </summary>
        /// <param name="glowType">The preset type</param>
        public void SetPresetGlow(ePresetExcelGlowType glowType)
        {
            Glow.Delete();
            if (glowType == ePresetExcelGlowType.None)
            {
                return;
            }

            var glowTypeString = glowType.ToString();
            var font = glowTypeString.Substring(0, glowTypeString.IndexOf('_'));
            var schemeColor = (eSchemeColor)Enum.Parse(typeof(eSchemeColor), font);
            Glow.Color.SetSchemeColor(schemeColor);
            Glow.Color.Transforms.AddAlpha(40);
            Glow.Color.Transforms.AddSaturationModulation(175);
            Glow.Radius = int.Parse(glowTypeString.Substring(font.Length + 1, glowTypeString.Length - font.Length - 3));
        }

        /// <summary>
        /// Set a predefined shadow matching the preset types in Excel
        /// </summary>
        /// <param name="shadowType">The preset type</param>
        public void SetPresetShadow(ePresetExcelShadowType shadowType)
        {
            InnerShadow.Delete();
            OuterShadow.Delete();
            PresetShadow.Delete();

            if (shadowType == ePresetExcelShadowType.None)
            {
                return;
            }

            if (shadowType <= ePresetExcelShadowType.InnerBottomRight)
            {
                SetPredefinedInnerShadow(shadowType);
            }
            else
            {
                SetPredefinedOuterShadow(shadowType);
            }
        }
        /// <summary>
        /// Set a predefined glow matching the preset types in Excel
        /// </summary>
        /// <param name="reflectionType">The preset type</param>
        public void SetPresetReflection(ePresetExcelReflectionType reflectionType)
        {
            Reflection.Delete();
            if (reflectionType == ePresetExcelReflectionType.TightTouching ||
               reflectionType == ePresetExcelReflectionType.Tight4Pt ||
               reflectionType == ePresetExcelReflectionType.Tight8Pt)
            {
                Reflection.Alignment = eRectangleAlignment.BottomLeft;
                Reflection.RotateWithShape = false;
                Reflection.Direction = 90;
                Reflection.VerticalScalingFactor = -100;
                Reflection.BlurRadius = 0.5;
                if (reflectionType == ePresetExcelReflectionType.TightTouching)
                {
                    Reflection.EndPosition = 35;
                    Reflection.StartOpacity = 52;
                    Reflection.EndOpacity = 0.3;
                    Reflection.Distance = 0;
                }
                else if (reflectionType == ePresetExcelReflectionType.Tight4Pt)
                {
                    Reflection.EndPosition = 38.5;
                    Reflection.StartOpacity = 50;
                    Reflection.EndOpacity = 0.3;
                    Reflection.Distance = 4;
                }
                else if (reflectionType == ePresetExcelReflectionType.Tight8Pt)
                {
                    Reflection.EndPosition = 40;
                    Reflection.StartOpacity = 50;
                    Reflection.EndOpacity = 0.275;
                    Reflection.Distance = 8;
                }
            }
            else if (reflectionType == ePresetExcelReflectionType.HalfTouching ||
                    reflectionType == ePresetExcelReflectionType.Half4Pt ||
                    reflectionType == ePresetExcelReflectionType.Half8Pt)
            {
                Reflection.Alignment = eRectangleAlignment.BottomLeft;
                Reflection.RotateWithShape = false;
                Reflection.Direction = 90;
                Reflection.VerticalScalingFactor = -100;
                Reflection.BlurRadius = 0.5;
                if (reflectionType == ePresetExcelReflectionType.HalfTouching)
                {
                    Reflection.EndPosition = 55;
                    Reflection.StartOpacity = 50;
                    Reflection.EndOpacity = 0.3;
                    Reflection.Distance = 0;
                }
                else if (reflectionType == ePresetExcelReflectionType.Half4Pt)
                {
                    Reflection.EndPosition = 55.5;
                    Reflection.StartOpacity = 50;
                    Reflection.EndOpacity = 0.300;
                    Reflection.Distance = 4;
                }
                else if (reflectionType == ePresetExcelReflectionType.Half8Pt)
                {
                    Reflection.EndPosition = 55.5;
                    Reflection.StartOpacity = 50;
                    Reflection.EndOpacity = 0.300;
                    Reflection.Distance = 8;
                }
            }
            else if (reflectionType == ePresetExcelReflectionType.FullTouching ||
                    reflectionType == ePresetExcelReflectionType.Full4Pt ||
                    reflectionType == ePresetExcelReflectionType.Full8Pt)
            {
                Reflection.Alignment = eRectangleAlignment.BottomLeft;
                Reflection.RotateWithShape = false;
                Reflection.Direction = 90;
                Reflection.VerticalScalingFactor = -100;
                Reflection.BlurRadius = 0.5;
                if (reflectionType == ePresetExcelReflectionType.FullTouching)
                {
                    Reflection.EndPosition = 90;
                    Reflection.StartOpacity = 50;
                    Reflection.EndOpacity = 0.3;
                    Reflection.Distance = 0;
                }
                else if (reflectionType == ePresetExcelReflectionType.Full4Pt)
                {
                    Reflection.EndPosition = 90;
                    Reflection.StartOpacity = 50;
                    Reflection.EndOpacity = 0.300;
                    Reflection.Distance = 4;
                }
                else if (reflectionType == ePresetExcelReflectionType.Full8Pt)
                {
                    Reflection.EndPosition = 92;
                    Reflection.StartOpacity = 50;
                    Reflection.EndOpacity = 0.295;
                    Reflection.Distance = 8;
                }
            }
        }
        #endregion
    }
}
