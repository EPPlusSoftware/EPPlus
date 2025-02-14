/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData
{
    /// <summary>
    /// Flags used for rich data.
    /// </summary>
    [Flags]
    internal enum RichValueKeyFlags
    {
        /// <summary>
        /// False indicates that we hide this key value pair (KVP) in the default Card View
        /// </summary>
        ShowInCardView = 0x01,
        /// <summary>
        /// False indicates that we hide this key value pair (KVP) from formulas and the object model
        /// </summary>
        ShowInDotNotation = 0x02,
        /// <summary>
        /// False indicates that we hide this key value pair (KVP) from AutoComplete, sort, filter, and Find
        /// </summary>
        ShowInAutoComplete = 0x04,
        /// <summary>
        /// True indicates that we do not write this key value pair (KVP) into the file, it only exists in memory
        /// </summary>
        ExcludeFromFile = 0x08,
        /// <summary>
        /// True indicates that we exclude this key value pair (KVP) when comparing rich values.
        /// </summary>
        ExcludeFromCalcComparison = 0x10,
    }
}
