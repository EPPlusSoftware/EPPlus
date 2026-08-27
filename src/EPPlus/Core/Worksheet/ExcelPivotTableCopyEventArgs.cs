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
namespace OfficeOpenXml.Core.Worksheet
{
    /// <summary>
    /// Provides context for a pivot table that is being copied to a new worksheet, and allows
    /// a custom name to be assigned to the copied pivot table.
    /// </summary>
    public class ExcelPivotTableCopyEventArgs
    {
        /// <summary>
        /// The name of the pivot table on the source worksheet.
        /// </summary>
        public string SourceTableName { get; internal set; }

        /// <summary>
        /// The name that was assigned to the copied pivot table by default, before this handler
        /// runs. When the worksheet is copied within the same workbook, this is a generated name
        /// (PivotTable1, PivotTable2, ...). When copied to another workbook, the original name is
        /// kept when it is still available, in which case this equals <see cref="SourceTableName"/>;
        /// if a pivot table with that name already exists in the target workbook, a generated name
        /// is used instead.
        /// </summary>
        public string DefaultName { get; internal set; }

        /// <summary>
        /// The name to assign to the copied pivot table. Leave as null to keep <see cref="DefaultName"/>.
        /// Setting this to an existing pivot table name will cause the same validation exception
        /// as a normal pivot table name assignment.
        /// </summary>
        public string NewName { get; set; }
    }
}