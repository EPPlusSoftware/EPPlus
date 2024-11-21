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
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect
{

    /// <summary>
    /// The outer shadow effect. A shadow is applied outside the edges of the drawing.
    /// </summary>
    public class ExcelDrawingOuterShadowEffect : ExcelDrawingInnerShadowEffect
    {
        private readonly string _shadowAlignmentPath = "{0}/@algn";
        private readonly string _rotateWithShapePath = "{0}/@rotWithShape";
        private readonly string _verticalSkewAnglePath = "{0}/@ky";
        private readonly string _horizontalSkewAnglePath = "{0}/@kx";
        private readonly string _verticalScalingFactorPath = "{0}/@sy";
        private readonly string _horizontalScalingFactorPath = "{0}/@sx";
        internal ExcelDrawingOuterShadowEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, schemaNodeOrder, path)
        {
            _shadowAlignmentPath = string.Format(_shadowAlignmentPath, path);
            _rotateWithShapePath = string.Format(_rotateWithShapePath, path);
            _verticalSkewAnglePath = string.Format(_verticalSkewAnglePath, path);
            _horizontalSkewAnglePath = string.Format(_horizontalSkewAnglePath, path);
            _verticalScalingFactorPath = string.Format(_verticalScalingFactorPath, path);
            _horizontalScalingFactorPath = string.Format(_horizontalScalingFactorPath, path);
        }
        /// <summary>
        /// The shadow alignment
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
        public double HorizontalSkewAngle
        {
            get
            {
                return  GetXmlNodeAngle(_horizontalSkewAnglePath);
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
        public double VerticalSkewAngle
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
        /// Horizontal scaling factor in percentage.
        /// A negative value causes a flip.
        /// </summary>
        public double HorizontalScalingFactor
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
        /// Vertical scaling factor in percentage.
        /// A negative value causes a flip.
        /// </summary>
        public double VerticalScalingFactor
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
    }
}