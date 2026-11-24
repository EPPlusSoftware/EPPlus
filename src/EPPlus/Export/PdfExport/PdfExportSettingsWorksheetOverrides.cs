/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
namespace OfficeOpenXml.Export.PdfExport
{
    /// <summary>
    /// Per-worksheet overrides for <see cref="PdfExportSettings"/>. A property left null inherits the
    /// document-level value; a non-null value overrides it for that worksheet only.
    /// </summary>
    public class PdfExportSettingsWorksheetOverrides
    {
        /// <summary>
        /// Per-worksheet bookmark prefix. When null (the default), the worksheet inherits
        /// <see cref="PdfExportSettings.BookmarkNamePrefix"/>. Set it to override the prefix for this worksheet.
        /// </summary>
        public string BookmarkNamePrefix { get; set; }
    }
}
