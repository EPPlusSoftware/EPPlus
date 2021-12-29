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
using OfficeOpenXml.Utils.Extensions;
using drawing =System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Represents a preset color
    /// </summary>
    public class ExcelDrawingPresetColor : XmlHelper
    {
        internal ExcelDrawingPresetColor(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        internal static ePresetColor GetPresetColor(drawing.Color presetColor)
        {
            return (ePresetColor)Enum.Parse(typeof(ePresetColor), TranslateFromColor(presetColor), true);
        }

        /// <summary>
        /// The preset color
        /// </summary>
        public ePresetColor Color
        {
            get
            {
                return GetXmlNodeString("@val").TranslatePresetColor();
            }
            set
            {
                SetXmlNodeString("@val", value.TranslateString());
            }
        }   

        private static string TranslateFromColor(drawing.Color c)
        {            
            if (c.IsEmpty || c.GetType().GetProperty(c.Name, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static) == null)
            {
                throw (new ArgumentException("A preset color cannot be set to empty or be an unnamed color"));
            }
            var s= c.Name.ToString();
            return s.Substring(0, 1).ToLower()+s.Substring(1);
        }

        internal const string NodeName = "a:prstClr";
    }
}