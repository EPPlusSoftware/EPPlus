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
    /// A preset shadow types
    /// </summary>
    public class ExcelDrawingPresetShadowEffect : ExcelDrawingShadowEffect
    {
        private readonly string _typePath = "{0}/@prst";
        internal ExcelDrawingPresetShadowEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, schemaNodeOrder, path)
        {
            _typePath = string.Format(_typePath, path);
        }
        /// <summary>
        /// The preset shadow type
        /// </summary>
        public ePresetShadowType Type
        {
            get
            {
                return GetXmlNodeString(_typePath).TranslatePresetShadowType();
            }
            set
            {
                SetXmlNodeString(_typePath, value.TranslateString());
                InitXml();
            }
        }
    }
}