/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/10/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Rounding method
    /// </summary>
    public enum RoundingMethod
    {
        /// <summary>
        /// Round decimal number to int using Convert.ToInt32
        /// </summary>
        Convert,
        /// <summary>
        /// Round decimal number to int using Math.Floor
        /// </summary>
        Floor
    }
}
