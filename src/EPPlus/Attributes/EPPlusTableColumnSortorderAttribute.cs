/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Attributes
{
    /// <summary>
    /// Attribute used by <see cref="ExcelRangeBase.LoadFromCollection{T}(IEnumerable{T})" /> to configure sorting of properties for the functions. Overrides any other configured sort order./>
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface, AllowMultiple = true)]
    public class EPPlusTableColumnSortOrderAttribute : Attribute
    {
        /// <summary>
        /// Property names used for the sort.
        /// </summary>
        public string[] Properties { get; set; }
    }
}
