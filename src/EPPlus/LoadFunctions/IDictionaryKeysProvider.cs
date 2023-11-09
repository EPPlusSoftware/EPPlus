/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  7/11/2023         EPPlus Software AB       EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    /// <summary>
    /// Provides keys of a property decorated with the <see cref="EPPlusDictionaryColumnAttribute"/>
    /// </summary>
    public interface IDictionaryKeysProvider
    {
        /// <summary>
        /// This function will return keys that will be used as column headers
        /// based on the <paramref name="key"/> that will be read from the <see cref="EPPlusDictionaryColumnAttribute"/>
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public IEnumerable<string> GetKeys(string key);
    }
}
