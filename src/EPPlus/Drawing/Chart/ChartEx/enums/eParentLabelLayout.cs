/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// The layout type for the parent labels
    /// </summary>
    public enum eParentLabelLayout
    {
        /// <summary>
        /// No parent labels are shown
        /// </summary>
        None,
        /// <summary>
        /// Parent label layout is a banner above the category
        /// </summary>
        Banner,
        /// <summary>
        /// Parent label is laid out within the category
        /// </summary>
        Overlapping
    }
}
