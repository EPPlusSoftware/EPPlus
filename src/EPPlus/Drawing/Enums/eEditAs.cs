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
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// How the drawing will be resized.
    /// </summary>
    public enum eEditAs
    {
        /// <summary>
        /// The Drawing is positioned absolute to the top left corner of the worksheet or chart and is NOT resized when rows and columns or the chart is resized. 
        /// </summary>
        Absolute,
        /// <summary>
        /// The Drawing will move with the worksheet but is NOT resized when rows and columns are resized. 
        /// This setting is not applicable for a drawing inside a chart.
        /// </summary>
        OneCell,
        /// <summary>
        /// The Drawing will move and resize when rows and columns are resized. 
        /// For drawings inside a chart the drawing will resize when the chart is resized.
        /// </summary>
        TwoCell,
        /// <summary>
        /// The Drawing will move and resize according to it's parent chart. 
        /// </summary>
        RelSize
    }
}