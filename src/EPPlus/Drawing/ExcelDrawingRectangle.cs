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
namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// The focus point for a non-liner gradient fill
    /// </summary>
    public class ExcelDrawingRectangle
    {

        internal ExcelDrawingRectangle(XmlHelper xml, string path, double defaultValue)
        {
            TopOffset = xml.GetXmlNodePercentage(path + "@t") ?? defaultValue;
            BottomOffset = xml.GetXmlNodePercentage(path + "@b") ?? defaultValue;
            LeftOffset = xml.GetXmlNodePercentage(path + "@l") ?? defaultValue;
            RightOffset = xml.GetXmlNodePercentage(path+"@r") ?? defaultValue;
        }

        internal ExcelDrawingRectangle(double defaultValue)
        {
            TopOffset = defaultValue;
            BottomOffset = defaultValue;
            LeftOffset = defaultValue;
            RightOffset = defaultValue;
        }
        /// <summary>
        /// Top offset in percentage
        /// </summary>
        public double TopOffset { get; set; }
        /// <summary>
        /// Bottom offset in percentage
        /// </summary>
        public double BottomOffset { get; set; }
        /// <summary>
        /// Left offset in percentage
        /// </summary>
        public double LeftOffset { get; set; }
        /// <summary>
        /// Right offset in percentage
        /// </summary>
        public double RightOffset { get; set; }
    }
}