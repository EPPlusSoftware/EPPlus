/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.Params
{
    /// <summary>
    /// Parameters for the LoadFromDictionaries method
    /// </summary>
    public class LoadFromDictionariesParams : LoadFunctionFunctionParamsBase
    {
        /// <summary>
        /// If set, only these keys will be included in the dataset
        /// </summary>
        public IEnumerable<string> Keys { get; private set; }

        /// <summary>
        /// The keys supplied to this function will be included in the dataset, all others will be ignored.
        /// </summary>
        /// <param name="keys">The keys to include</param>
        public void SetKeys(params string[] keys)
        {
            Keys = keys;
        }

        /// <summary>
        /// Sets how headers should be parsed before added to the worksheet, see <see cref="HeaderParsingTypes"/>
        /// </summary>
        public HeaderParsingTypes HeaderParsingType { get; set; } = HeaderParsingTypes.UnderscoreToSpace;

        /// <summary>
        /// Data types used when setting data in the spreadsheet range (defined from left to right per column).
        /// </summary>
        public eDataTypes[] DataTypes { get; set; }
    }
}
