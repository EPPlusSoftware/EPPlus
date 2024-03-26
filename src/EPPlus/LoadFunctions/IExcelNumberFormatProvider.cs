/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  17/2/2024         EPPlus Software AB       EPPlus 7.1
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    /// <summary>
    /// This interface is used to provide number formats for ranges
    /// in runtime. The number formats are mapped to a number that can be used
    /// to refer to a specific format.
    /// </summary>
    public interface IExcelNumberFormatProvider
    {
        /// <summary>
        /// Returns a number format by its <paramref name="numberFormatId"/>
        /// </summary>
        /// <param name="numberFormatId"></param>
        /// <returns>A number format that can be used on <see cref="ExcelRangeBase">ranges</see></returns>
        public string GetFormat(int numberFormatId);
    }
}
