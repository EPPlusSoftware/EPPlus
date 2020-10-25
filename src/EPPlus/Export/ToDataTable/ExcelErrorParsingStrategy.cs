/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    public enum ExcelErrorParsingStrategy
    {
        /// <summary>
        /// Excel Errors in cells will be handles as blank cells
        /// </summary>
        HandleExcelErrorsAsBlankCells,
        /// <summary>
        /// An exception will be thrown when an error occurs in a cell
        /// </summary>
        ThrowException,
        /// <summary>
        /// If an error is detected, the entire row will be ignored
        /// </summary>
        IgnoreRowWithErrors
    }
}
