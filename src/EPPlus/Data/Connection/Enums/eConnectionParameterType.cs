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
namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Connection parameter types.
    /// </summary>
    public enum eConnectionParameterType
    {
        /// <summary>
        /// Prompt the user on each refresh for a parameter value
        /// </summary>
        Prompt,
        /// <summary>
        /// Use a constant value on each refresh for the parameter value.
        /// </summary>
        Value,
        /// <summary>
        /// Get the parameter value from a cell on each refresh
        /// </summary>
        Cell
    }
}