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
    /// Preset soft edges types in Excel
    /// </summary>
    public enum ePresetExcelSoftEdgesType
    {
        /// <summary>
        /// No soft edges
        /// </summary>
        None,
        /// <summary>
        /// Soft edges 1pt
        /// </summary>
        SoftEdge1Pt,
        /// <summary>
        /// Soft edges 2.5pt
        /// </summary>
        SoftEdge2_5Pt,
        /// <summary>
        /// Soft edges 5pt
        /// </summary>
        SoftEdge5Pt,
        /// <summary>
        /// Soft edges 10pt
        /// </summary>
        SoftEdge10Pt,
        /// <summary>
        /// Soft edges 25pt
        /// </summary>
        SoftEdge25Pt,
        /// <summary>
        /// Soft edges 50pt
        /// </summary>
        SoftEdge50Pt
    }
}