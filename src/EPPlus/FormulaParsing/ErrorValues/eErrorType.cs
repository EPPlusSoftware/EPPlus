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
namespace OfficeOpenXml
{
    /// <summary>
    /// Represents the errortypes in excel
    /// </summary>
    public enum eErrorType
    {
        /// <summary>
        /// Division by zero
        /// </summary>
        Div0,
        /// <summary>
        /// Not applicable
        /// </summary>
        NA,
        /// <summary>
        /// Name error
        /// </summary>
        Name,
        /// <summary>
        /// Null error
        /// </summary>
        Null,
        /// <summary>
        /// Num error
        /// </summary>
        Num,
        /// <summary>
        /// Reference error
        /// </summary>
        Ref,
        /// <summary>
        /// Value error
        /// </summary>
        Value,
        /// <summary>
        /// Calc error
        /// </summary>
        Calc,
        /// <summary>
        /// Spill error from a dynamic array formula.
        /// </summary>
        Spill,
    }
}
