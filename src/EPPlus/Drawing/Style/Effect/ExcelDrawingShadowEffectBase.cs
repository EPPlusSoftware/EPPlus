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
    /// Base class for shadow effects
    /// </summary>
    public abstract class ExcelDrawingShadowEffectBase : ExcelDrawingEffectBase 
    {
        private readonly string _distancePath = "{0}/@dist";

        internal ExcelDrawingShadowEffectBase(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, schemaNodeOrder, path)
        {
            _distancePath = string.Format(_distancePath, path);
        }
        
        /// <summary>
        /// How far to offset the shadow is in pixels
        /// </summary>
        public double Distance
        {
            get
            {
                return GetXmlNodeEmuToPt(_distancePath);
            }
            set
            {
                SetXmlNodeEmuToPt(_distancePath, value);
            }
        }        
    }
}