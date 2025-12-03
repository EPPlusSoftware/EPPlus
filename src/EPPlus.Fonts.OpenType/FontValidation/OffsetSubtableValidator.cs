/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.FontValidation
{
    internal class OffsetSubtableValidator
    {
        public FontValidationSeverity LogLevel { get; set; } = FontValidationSeverity.All;

        public TableValidationResult Validate(OpenTypeFont font, FontValidationContext context, FontValidationSeverity logLevel)
        {
            var result = new TableValidationResult { TableName = "OffsetSubtable", LogLevel = logLevel };

            // Läs värden
            ushort numTables = font.NumTables;
            ushort searchRange = font.SearchRange;
            ushort entrySelector = font.EntrySelector;
            ushort rangeShift = font.RangeShift;

            // Beräkna förväntade värden
            int maxPower = (int)Math.Pow(2, (int)Math.Floor(Math.Log(numTables, 2)));
            ushort expectedSearchRange = (ushort)(maxPower * 16);
            ushort expectedEntrySelector = (ushort)Math.Floor(Math.Log(numTables, 2));
            ushort expectedRangeShift = (ushort)((numTables * 16) - expectedSearchRange);

            // Validera
            if (searchRange != expectedSearchRange)
                result.AddMessage(FontValidationSeverity.Error, $"searchRange mismatch: expected {expectedSearchRange}, got {searchRange}.");
            if (entrySelector != expectedEntrySelector)
                result.AddMessage(FontValidationSeverity.Error, $"entrySelector mismatch: expected {expectedEntrySelector}, got {entrySelector}.");
            if (rangeShift != expectedRangeShift)
                result.AddMessage(FontValidationSeverity.Error, $"rangeShift mismatch: expected {expectedRangeShift}, got {rangeShift}.");

            return result;
        }
    }
}
