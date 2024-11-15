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
namespace OfficeOpenXml.Drawing.Style.Fill
{

    /// <summary>
    /// Settings specific for linear gradiant fills
    /// </summary>
    public class ExcelDrawingGradientFillLinearSettings
    {
        internal ExcelDrawingGradientFillLinearSettings()
        {
        }
        internal ExcelDrawingGradientFillLinearSettings(XmlHelper xml)
        {
            Angel = xml.GetXmlNodeAngle("a:lin/@ang");
            Scaled = xml.GetXmlNodeBool("a:lin/@scaled", false);
        }

        /// <summary>
        /// The direction of color change for the gradient.To define this angle, let its value
        /// be x measured clockwise.Then( -sin x, cos x) is a vector parallel to the line of constant color in the gradient fill.
        /// </summary>
        public double Angel { get; set; }
        /// <summary>
        /// If the gradient angle scales with the fill.
        /// </summary>
        public bool Scaled { get; set; }       
    }
}