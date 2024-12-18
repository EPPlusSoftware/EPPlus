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
using OfficeOpenXml.Drawing.Style;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect
{

    /// <summary>
    /// The reflection effect
    /// </summary>
    public class ExcelDrawingReflectionEffect  : ExcelDrawingShadowEffectBase
    {
        private readonly string _directionPath = "{0}/@dir";
        private readonly string _startPositionPath = "{0}/@stPos";
        private readonly string _startOpacityPath = "{0}/@stA";
        private readonly string _endPositionPath = "{0}/@endPos";
        private readonly string _endOpacityPath = "{0}/@endA";
        private readonly string _fadeDirectionPath = "{0}/@fadeDir";
        private readonly string _shadowAlignmentPath = "{0}/@algn";
        private readonly string _rotateWithShapePath = "{0}/@rotWithShape";
        private readonly string _verticalSkewAnglePath = "{0}/@ky";
        private readonly string _horizontalSkewAnglePath = "{0}/@kx";
        private readonly string _verticalScalingFactorPath = "{0}/@sy";
        private readonly string _horizontalScalingFactorPath = "{0}/@sx";
        private readonly string _blurRadPath = "{0}/@blurRad";
        internal ExcelDrawingReflectionEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, schemaNodeOrder, path)
        {
            _startPositionPath = string.Format(_startPositionPath, path);
            _startOpacityPath = string.Format(_startOpacityPath, path);
            _endPositionPath = string.Format(_endPositionPath, path);
            _endOpacityPath = string.Format(_endOpacityPath, path);
            _fadeDirectionPath = string.Format(_fadeDirectionPath, path);
            _shadowAlignmentPath = string.Format(_shadowAlignmentPath, path);
            _rotateWithShapePath = string.Format(_rotateWithShapePath, path);
            _verticalSkewAnglePath = string.Format(_verticalSkewAnglePath, path);
            _horizontalSkewAnglePath = string.Format(_horizontalSkewAnglePath, path);
            _verticalScalingFactorPath = string.Format(_verticalScalingFactorPath, path);
            _horizontalScalingFactorPath = string.Format(_horizontalScalingFactorPath, path);
            _directionPath = string.Format(_directionPath, path);
            _blurRadPath = string.Format(_blurRadPath, path);
        }
        /// <summary>
        /// The start position along the alpha gradient ramp of the alpha value.
        /// </summary>
        public double? StartPosition
        {
            get
            {
                return GetXmlNodePercentage(_startPositionPath) ?? 0;
            }
            set
            {
                SetXmlNodePercentage(_startPositionPath, value, false);
            }
        }
        /// <summary>
        /// The starting reflection opacity
        /// </summary>
        public double? StartOpacity
        {
            get
            {
                return GetXmlNodePercentage(_startOpacityPath) ?? 100;
            }
            set
            {
                SetXmlNodePercentage(_startOpacityPath, value, false);
            }
        }

        /// <summary>
        /// The end position along the alpha gradient ramp of the alpha value.
        /// </summary>
        public double? EndPosition
        {
            get
            {
                return GetXmlNodePercentage(_endPositionPath) ?? 100;
            }
            set
            {
                SetXmlNodePercentage(_endPositionPath, value, false);
            }
        }
        /// <summary>
        /// The ending reflection opacity
        /// </summary>
        public double? EndOpacity
        {
            get
            {
                return GetXmlNodePercentage(_endOpacityPath) ?? 0;
            }
            set
            {
                SetXmlNodePercentage(_endOpacityPath, value, false);
            }
        }
        /// <summary>
        /// The direction to offset the reflection
        /// </summary>
        public double? FadeDirection
        {
            get
            {
                return GetXmlNodeAngle(_fadeDirectionPath, 90);
            }
            set
            {
                SetXmlNodeAngle(_fadeDirectionPath, value, "FadeDirection", -90, 90);
            }
        }
        /// <summary>
        /// Alignment
        /// </summary>
        public eRectangleAlignment Alignment
        {
            get
            {
                return GetXmlNodeString(_shadowAlignmentPath).TranslateRectangleAlignment();
            }
            set
            {
                if (value == eRectangleAlignment.Bottom)
                {
                    DeleteNode(_shadowAlignmentPath);
                }
                else
                {
                    SetXmlNodeString(_shadowAlignmentPath, value.TranslateString());
                }
            }
        }
        /// <summary>
        /// If the shadow rotates with the shape
        /// </summary>
        public bool RotateWithShape
        {
            get
            {
                return GetXmlNodeBool(_rotateWithShapePath, true);
            }
            set
            {
                SetXmlNodeBool(_rotateWithShapePath, value, true);
            }
        }
        /// <summary>
        /// Horizontal skew angle.
        /// Ranges from -90 to 90 degrees 
        /// </summary>
        public double? HorizontalSkewAngle
        {
            get
            {
                return GetXmlNodeAngle(_horizontalSkewAnglePath);
            }
            set
            {
                SetXmlNodeAngle(_horizontalSkewAnglePath, value, "HorizontalSkewAngle", -90, 90);
            }
        }
        /// <summary>
        /// Vertical skew angle.
        /// Ranges from -90 to 90 degrees 
        /// </summary>
        public double? VerticalSkewAngle
        {
            get
            {
                return GetXmlNodeAngle(_verticalSkewAnglePath);
            }
            set
            {
                SetXmlNodeAngle(_verticalSkewAnglePath, value, "HorizontalSkewAngle", -90, 90);
            }
        }
        /// <summary>
        /// Horizontal scaling factor in percentage .
        /// A negative value causes a flip.
        /// </summary>
        public double? HorizontalScalingFactor
        {
            get
            {
                return GetXmlNodePercentage(_horizontalScalingFactorPath) ?? 100;
            }
            set
            {
                SetXmlNodePercentage(_horizontalScalingFactorPath, value, true, 10000);
            }
        }
        /// <summary>
        /// Vertical scaling factor in percentage .
        /// A negative value causes a flip.
        /// </summary>
        public double? VerticalScalingFactor
        {
            get
            {
                return GetXmlNodePercentage(_verticalScalingFactorPath) ?? 100;
            }
            set
            {
                SetXmlNodePercentage(_verticalScalingFactorPath, value, true, 10000);
            }
        }
        /// <summary>
        /// The direction to offset the shadow
        /// </summary>
        public double? Direction
        {
            get
            {
                return GetXmlNodeAngle(_directionPath);
            }
            set
            {
                SetXmlNodeAngle(_directionPath, value, "Direction");
            }
        }
        /// <summary>
        /// The blur radius.
        /// </summary>
        public double? BlurRadius
        {
            get
            {
                return GetXmlNodeEmuToPt(_blurRadPath);
            }
            set
            {
                SetXmlNodeEmuToPt(_blurRadPath, value);
            }
        }
    }
}