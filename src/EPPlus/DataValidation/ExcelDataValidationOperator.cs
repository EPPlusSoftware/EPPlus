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

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Operator for comparison between Formula and Formula2 in a validation.
    /// </summary>
    public enum ExcelDataValidationOperator
    {
        /// <summary>
        /// The value of the validated cell should be between two values
        /// </summary>
        between = 0,
        /// <summary>
        /// The value of the validated cell should be eqal to a specific value
        /// </summary>
        equal = 2,
        /// <summary>
        /// The value of the validated cell should be greater than a specific value
        /// </summary>
        greaterThan = 6,
        /// <summary>
        /// The value of the validated cell should be greater than or equal to a specific value
        /// </summary>
        greaterThanOrEqual = 7,
        /// <summary>
        /// The value of the validated cell should be less than a specific value
        /// </summary>
        lessThan = 4,
        /// <summary>
        /// The value of the validated cell should be less than or equal to a specific value
        /// </summary>
        lessThanOrEqual = 5,
        /// <summary>
        /// The value of the validated cell should not be between two specified values
        /// </summary>
        notBetween = 1,
        /// <summary>
        /// The value of the validated cell should not be eqal to a specific value
        /// </summary>
        notEqual = 3
    }
}
