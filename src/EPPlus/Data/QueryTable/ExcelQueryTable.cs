/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using OfficeOpenXml.Core;
using OfficeOpenXml.Data.Connection;
using OfficeOpenXml.Table;

namespace OfficeOpenXml.Data.QueryTable
{
    /// <summary>
    /// A Query Table connected to an <see cref="ExcelTable" /> object.
    /// </summary>
    public class ExcelQueryTable : DocumentPart<ExcelQueryTable>
    {
        internal ExcelQueryTable(IDocumentPart<ExcelQueryTable> dp) : base(dp)
        {
        }
        /// <summary>
        /// The fields mapping to the column in the table.
        /// </summary>
        public EPPlusReadOnlyList<ExcelQueryTableField> Fields { get; } = new EPPlusReadOnlyList<ExcelQueryTableField>();
        /// <summary>
        /// A collection of deleted fields from the query table.
        /// </summary>
        public EPPlusReadOnlyList<string> DeletedFields { get; } = new EPPlusReadOnlyList<string>();
        /// <summary>
        /// Specifies whether to automatically adjust column widths on refresh to fit the data retrieved. true if column widths should be adjusted.
        /// </summary>
        public bool AdjustColumnWidth { get; set; }
        /// <summary>
        /// If true, apply legacy table autoformat alignment properties.
        ///The possible values for this attribute are defined by the W3C XML Schema boolean datatype.
        /// </summary>
        public bool? ApplyAlignmentFormats { get; set; }
        /// <summary>
        /// If true, apply legacy table autoformat border properties.
        /// </summary>
        public bool? ApplyBorderFormats { get; set; }

        /// <summary>
        /// If true, apply legacy table autoformat font properties.
        /// </summary>
        public bool? ApplyFontFormats { get; set; }

        /// <summary>
        /// If true, apply legacy table autoformat number format properties.
        /// </summary>
        public bool? ApplyNumberFormats { get; set; }

        /// <summary>
        /// If true, apply legacy table autoformat pattern properties.
        /// </summary>
        public bool? ApplyPatternFormats { get; set; }

        /// <summary>
        /// If true, apply legacy table autoformat width/height properties.
        /// </summary>
        public bool? ApplyWidthHeightFormats { get; set; }

        /// <summary>
        /// Identifies which legacy table autoformat to apply.
        /// </summary>
        public int? AutoFormatId { get; set; }

        /// <summary>
        /// Specifies whether the query table shall try to refresh data in the background.
        /// </summary>
        public bool? BackgroundRefresh { get; set; }

        /// <summary>
        /// Specifies the ID number of the external data connection to use to refresh data in the query table.
        /// </summary>
        internal int ConnectionId { get; set; }
        /// <summary>
        /// The connection used for the query table.
        /// </summary>
        public ExcelConnection Connection { get; internal set; }
        /// <summary>
        /// Specifies whether the connection used with this query table shall be editable.
        /// If true, then the connection is not editable.
        /// </summary>
        public bool DisableEdit { get; set; }

        /// <summary>
        /// Specifies whether the query table shall be refreshable.
        /// If true, then the query table is not refreshable.
        /// </summary>
        public bool DisableRefresh { get; set; }

        /// <summary>
        /// Specifies whether formulas in columns adjacent to the query table should be filled down when refreshed.
        /// </summary>
        public bool FillFormulas { get; set; }

        /// <summary>
        /// Specifies whether the first background data refresh has completed.
        /// If true, the very first background refresh had not completed when the file was saved.
        /// </summary>
        public bool? FirstBackgroundRefresh { get; set; }

        /// <summary>
        /// Specifies how to handle variable numbers of rows between refresh operations.
        /// </summary>
        public QueryTableGrowShrinkType GrowShrinkType { get; set; }

        /// <summary>
        /// Specifies whether the query table has a first row with column titles.
        /// </summary>
        public bool Headers { get; set; }

        /// <summary>
        /// Specifies whether this query table is in an intermediate state.
        /// </summary>
        public bool Intermediate { get; set; }

        /// <summary>
        /// Specifies the name of the query table. Reqired.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Specifies whether formatting in the query table should be preserved and copied to new rows.
        /// </summary>
        public bool PreserveFormatting { get; set; }

        /// <summary>
        /// Specifies whether the query table shall refresh automatically when the document is loaded.
        /// </summary>
        public bool RefreshOnLoad { get; set; }

        /// <summary>
        /// Specifies whether all data shall be removed before saving the document.
        /// </summary>
        public bool? RemoveDataOnSave { get; set; }

        /// <summary>
        /// Specifies whether the query table shall include a first column of row numbers.
        /// </summary>
        public bool RowNumbers { get; set; }
        /// <summary>
        /// The destination range for the query table.
        /// </summary>
        public ExcelAddressBase DestinationRange { get; internal set; }
    }
}        
