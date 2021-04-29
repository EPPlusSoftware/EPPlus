/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
namespace OfficeOpenXml
{
    public enum eDdeValueType
    {
        /// <summary>
        /// The value is a boolean.
        /// </summary>
        Boolean,
        /// <summary>
        /// The value is an error.
        /// </summary>
        Error,
        /// <summary>
        /// The value is a real number.
        /// </summary>
        Number,
        /// <summary>
        /// The value is nil.
        /// </summary>
        Nil,
        /// <summary>
        /// The value is a string.
        /// </summary>
        String
    }
}