/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  1/4/2021         EPPlus Software AB           EPPlus Interfaces 1.0
 *************************************************************************************************/

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// Interface for measuring width and height of texts.
    /// </summary>
    public interface ITextMeasurer
    {
        /// <summary>
        /// Should return true if the text measurer is valid for this environment. 
        /// </summary>
        /// <returns>True if the measurer can be used else false.</returns>
        bool ValidForEnvironment();
        /// <summary>
        /// Measures width and height of the parameter <paramref name="text"/>.
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="font">The <see cref="ExcelFont">font</see> to measure</param>
        /// <returns></returns>
        TextMeasurement MeasureText(string text, ExcelFont font);
    }
}
