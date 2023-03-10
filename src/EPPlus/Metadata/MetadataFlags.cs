/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Metadata
{
    internal partial class ExcelMetadataType
    {
        [Flags]
        internal enum MetadataFlags
        {
            GhostRow = 1,
            GhostCol = 1 << 1,
            Edit = 1 << 2,
            Delete = 1 << 3,
            Copy = 1 << 4,
            PasteAll = 1 << 5,
            PasteFormulas = 1 << 6,
            PasteValues = 1 << 7,
            PasteFormats = 1 << 8,
            PasteComments = 1 << 9,
            PasteDataValidation = 1 << 10,
            PasteBorders = 1 << 10,
            PasteColWidths = 1 << 11,
            PasteNumberFormats = 1 << 12,
            Merge = 1 << 13,
            SplitFirst = 1 << 14,
            SplitAll = 1 << 15,
            RowColShift = 1 << 16,
            ClearAll = 1 << 17,
            ClearFormats = 1 << 18,
            ClearContents = 1 << 19,
            ClearComments = 1 << 20,
            Assign = 1 << 21,
            Coerce = 1 << 22,
            Adjust = 1 << 23,
            CellMeta = 1 << 24,
        }
    }
}