/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class StyleCache : Dictionary<string, int>
    {
        internal StyleCache()
        {
        }

        /// <summary>
        /// Returns true if key already existed
        /// Returns false if key could not be found and had to be created.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        internal bool GetOrCreateId(string key, out int id)
        {
            if (ContainsKey(key))
            {
                id = base[key];
                return true;
            }
            else
            {
                id = Count + 1;
                Add(key, id);
                return false;
            }
        }
    }
}
