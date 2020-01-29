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
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect
{

    /// <summary>
    /// The inner shadow effect. A shadow is applied within the edges of the drawing.
    /// </summary>
    public class ExcelDrawingInnerShadowEffect : ExcelDrawingShadowEffect
    {
        private readonly string _blurRadPath = "{0}/@blurRad";

        internal ExcelDrawingInnerShadowEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, schemaNodeOrder, path)
        {
            _blurRadPath = string.Format(_blurRadPath, path);
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
                InitXml();
            }
        }
    }
}