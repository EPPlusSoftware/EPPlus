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
namespace OfficeOpenXml.Table.PivotTable
{
    
    /// <summary>
    /// Defines the axis for a PivotTable
    /// </summary>
    public enum ePivotFieldAxis
    {
        /// <summary>
        /// None
        /// </summary>
        None=-1,
        /// <summary>
        /// Column axis
        /// </summary>
        Column,
        /// <summary>
        /// Page axis (Include Count Filter) 
        /// 
        /// </summary>
        Page,
        /// <summary>
        /// Row axis
        /// </summary>
        Row,
        /// <summary>
        /// Values axis
        /// </summary>
        Values 
    }
}