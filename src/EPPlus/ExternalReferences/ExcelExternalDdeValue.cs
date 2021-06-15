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
using System.Collections.Generic;

namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// Represents a value for a DDE item.
    /// </summary>
    public class ExcelExternalDdeValue
    {
        /// <summary>
        /// The data type of the value
        /// </summary>
        public eDdeValueType DdeValueType { get; internal set; } = eDdeValueType.Number;
        /// <summary>
        /// The value of the item
        /// </summary>
        public string Value { get; internal set; }
    }
}