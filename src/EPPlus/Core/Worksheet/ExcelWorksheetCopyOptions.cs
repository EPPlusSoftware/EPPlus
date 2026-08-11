/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                   Change
*************************************************************************************************
 08/05/2026         EPPlus Software AB       Added
*************************************************************************************************/
using System;

namespace OfficeOpenXml.Core.Worksheet
{
    /// <summary>
    /// Used to specify options when copying a worksheet.
    /// </summary>
    public class ExcelWorksheetCopyOptions
    {
        internal static ExcelWorksheetCopyOptions Default => new ExcelWorksheetCopyOptions();

        /// <summary>
        /// A handler that is invoked for each table that is copied to the new worksheet.
        /// Use this to assign a custom name to the copied table. When a worksheet is copied
        /// within the same workbook, copied tables are otherwise given a generated name
        /// (Table1, Table2, ...). Set <see cref="ExcelTableCopyEventArgs.NewName"/> on the
        /// argument to rename the copied table. The rename is applied through the same path
        /// as a normal <see cref="OfficeOpenXml.Table.ExcelTable.Name"/> assignment, so
        /// formula references are updated and name uniqueness is validated.
        /// </summary>
        public Action<ExcelTableCopyEventArgs> TableCopyHandler { get; set; }
    }
}
