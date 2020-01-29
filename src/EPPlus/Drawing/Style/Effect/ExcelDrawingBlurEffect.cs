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
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect
{
    /// <summary>
    /// A blur effect that is applied to the shape, including its fill
    /// </summary>
    public class ExcelDrawingBlurEffect : ExcelDrawingEffectBase
    {
        private readonly string _radiusPath = "{0}/@rad";
        private readonly string _glowBoundsPath = "{0}/@grow";
        internal ExcelDrawingBlurEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, schemaNodeOrder, path)
        {
            _radiusPath = string.Format(_radiusPath, path);
            _glowBoundsPath = string.Format(_glowBoundsPath, path);
        }
        /// <summary>
        /// The radius of blur in points
        /// </summary>
        public double? Radius
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_radiusPath) ?? 0;
            }
            set
            {
                SetXmlNodeEmuToPt(_radiusPath, value);
            }
        }
        /// <summary>
        /// If the bounds of the object will be grown as a result of the blurring.
        /// Default is true
        /// </summary>
        public bool GrowBounds
        {
            get
            {
                return GetXmlNodeBool(_glowBoundsPath, true);
            }
            set
            {
                SetXmlNodeBool(_glowBoundsPath, value, true);
            }
        }

    }
}
