/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// The power query meta data item type
    /// </summary>
    public enum ePowerQueryMetadataItemType
    {
        /// <summary>
        /// The item applies to all formulas
        /// </summary>
        AllFormulas,
        /// <summary>
        /// The item applies to the formula specified in the ItemPath.
        /// </summary>
        Formula
    }
}