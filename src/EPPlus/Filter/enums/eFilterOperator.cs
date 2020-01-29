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
namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Operator used by the filter comparison
    /// </summary>
    public enum eFilterOperator
    {
        /// <summary>
        /// Show results which are equal to the criteria
        /// </summary>
        Equal,
        /// <summary>
        /// Show results which are greater than the criteria
        /// </summary>
        GreaterThan,
        /// <summary>
        /// Show results which are greater than or equal to the criteria
        /// </summary>
        GreaterThanOrEqual,
        /// <summary>
        /// Show results which are less than the criteria
        /// </summary>
        LessThan,
        /// <summary>
        /// Show results which are less than or equal to the criteria
        /// </summary>
        LessThanOrEqual,
        /// <summary>
        /// Show results which are Not Equal to the criteria
        /// </summary>
        NotEqual
    }
}