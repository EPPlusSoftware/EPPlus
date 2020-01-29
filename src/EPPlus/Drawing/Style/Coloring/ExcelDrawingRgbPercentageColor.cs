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
namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// A color using the red, green, blue RGB color model.
    /// Each component, red, green, and blue is expressed as a percentage from 0% to 100%.
    /// A linear gamma of 1.0 is assumed
    /// </summary>
    public class ExcelDrawingRgbPercentageColor : XmlHelper
    {
        internal ExcelDrawingRgbPercentageColor(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        /// <summary>
        /// The percentage of red.
        /// </summary>
        public double RedPercentage 
        {
            get
            {
                return (double)GetXmlNodePercentage("@r");
            }
            set
            {
                SetXmlNodePercentage("@r", value, false);
                
            }
        }
        /// <summary>
        /// The percentage of green.
        /// </summary>
        public double GreenPercentage
        {
            get
            {
                return (double)GetXmlNodePercentage("@g");
            }
            set
            {
                SetXmlNodePercentage("@g", value, false);
            }
        }
        /// <summary>
        /// The percentage of blue.
        /// </summary>
        public double BluePercentage
        {
            get
            {
                return (double)GetXmlNodePercentage("@b");
            }
            set
            {
                SetXmlNodePercentage("@b", value, false);
            }
        }
        internal const string NodeName = "a:scrgbClr";
    }
}