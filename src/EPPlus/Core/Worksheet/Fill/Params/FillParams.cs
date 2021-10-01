/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    08/11/2021         EPPlus Software AB       EPPlus 5.8
 *************************************************************************************************/
namespace OfficeOpenXml.Core.Worksheet.Fill
{
    public class FillParams
    {
        /// <summary>
        /// If the fill starts from the top-left cell or the bottom right cell.
        /// </summary>
        public eFillStartPosition StartPosition { get; set; } = eFillStartPosition.TopLeft;
        /// <summary>
        /// The direction of the fill
        /// </summary>
        public eFillDirection Direction { get; set; } = eFillDirection.Column;
        /// <summary>
        /// The number format to be appled to the range.
        /// </summary>
        public string NumberFormat { get; set; } = null;
    }
}