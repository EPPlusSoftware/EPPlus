/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/286/2021         EPPlus Software AB       EPPlus 5.7.5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Attributes
{
    /// <summary>
    /// Attribute used by <see cref="ExcelRangeBase.LoadFromCollection{T}(IEnumerable{T})" /> to support complex type properties/>
    /// </summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property | AttributeTargets.Field)]    
    public class EpplusNestedTableColumnAttribute : Attribute
    {
        /// <summary>
        /// Order of the columns value, default value is 0
        /// </summary>
        public int Order
        {
            get;
            set;
        }

        /// <summary>
        /// This will prefix all names derived by members in the complex type.
        /// </summary>
        public string HeaderPrefix
        {
            get;
            set;
        }
    }
}
