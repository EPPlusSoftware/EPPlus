﻿/*************************************************************************************************
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
using OfficeOpenXml.Core;

namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// A collection of <see cref="ExcelExternalDdeValue" />
    /// </summary>
    public class ExcelExternalDdeValueCollection : EPPlusReadOnlyList<ExcelExternalDdeValue>
    {
        /// <summary>
        /// The number of rows returned by the server for this dde item.
        /// </summary>
        public int Rows { get; set; }
        /// <summary>
        /// The number of columns returned by the server for this dde item.
        /// </summary>
        public int Columns { get; set; }
    }
}