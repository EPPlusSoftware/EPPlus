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
using EPPlus.Fonts.OpenType.FontValidation;

namespace EPPlus.Fonts.OpenType.Tables.Os2
{
    internal class Os2TableValidator : TableValidatorBase<Os2Table>
    {
        public override string TableName
        {
            get { return TableNames.Os2; }
        }

        public override Type TableType => typeof(Os2Table);



        public override TableValidationResult Validate(Os2Table table, FontValidationContext context)
        {
            var result = new TableValidationResult { TableName = TableName };

            // -------------------------
            // Basic OS/2 checks
            // -------------------------

            // Version check
            if (table.version < 0 || table.version > 5)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Unexpected OS/2 version: {table.version}. Expected 0–5.");
            }

            // usWeightClass (100–900)
            if (table.usWeightClass < 100 || table.usWeightClass > 900)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"usWeightClass out of range: {table.usWeightClass}. Expected 100–900.");
            }

            // usWidthClass (1–9)
            if (table.usWidthClass < 1 || table.usWidthClass > 9)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"usWidthClass out of range: {table.usWidthClass}. Expected 1–9.");
            }

            // fsType basic info
            if ((table.fsType & 0x0002) != 0)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    "Font has restricted embedding (fsType bit 1 set).");
            }

            // Subscript/Superscript sizes > 0
            if (table.ySubscriptXSize <= 0 || table.ySuperscriptXSize <= 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    "Subscript/Superscript sizes should be greater than zero.");
            }

            // Panose check (warn if all zeros)
            if (table.panose != null && table.panose.Length == 10 && Array.TrueForAll(table.panose, b => b == 0))
            {
                result.AddMessage(FontValidationSeverity.Warning, "Panose classification is all zeros.");
            }

            // -------------------------
            // Cross-table: OS/2 vs cmap
            // -------------------------
            if (context.Font.CmapTable != null)
            {
                int minChar = context.Font.CmapTable.GetMinCharCode();
                int maxChar = context.Font.CmapTable.GetMaxCharCode();
                if (minChar < table.usFirstCharIndex || maxChar > table.usLastCharIndex)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Cmap character range ({minChar}-{maxChar}) is outside OS/2 declared range ({table.usFirstCharIndex}-{table.usLastCharIndex}).");
                }

                // Check default and break char exist in cmap
                if (!context.Font.CmapTable.ContainsChar(table.usDefaultChar))
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"usDefaultChar ({table.usDefaultChar}) not found in cmap.");
                }
                if (!context.Font.CmapTable.ContainsChar(table.usBreakChar))
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"usBreakChar ({table.usBreakChar}) not found in cmap.");
                }
            }

            // -------------------------
            // Critical rules for subsetting
            // -------------------------

            // Embedding permissions
            if ((table.fsType & 0x0002) != 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    "Embedding is restricted (fsType bit 1 set). Subsetting cannot proceed.");
            }
            if ((table.fsType & 0x0008) != 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    "No subsetting allowed (fsType bit 3 set).");
            }
            if ((table.fsType & 0x0004) != 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    "Preview & Print embedding only (fsType bit 2 set). Check usage context.");
            }

            // Metrics must be valid
            if (table.sTypoAscender == 0 || table.sTypoDescender == 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    "sTypoAscender and sTypoDescender must be non-zero for proper line spacing.");
            }
            if (table.usWinAscent == 0 || table.usWinDescent == 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    "usWinAscent and usWinDescent must be non-zero for bounding box calculations.");
            }

            // -------------------------
            // Cross-table: OS/2 vs head
            // -------------------------
            if (context.Font.HeadTable != null)
            {
                if (table.usWinAscent < context.Font.HeadTable.Ymax || table.usWinDescent < Math.Abs(context.Font.HeadTable.Ymin))
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        "usWinAscent/usWinDescent do not cover font bounding box from head table. Consider updating after subsetting.");
                }
            }

            // -------------------------
            // Cross-table: OS/2 vs hhea
            // -------------------------
            if (context.Font.HheaTable != null)
            {
                if (table.sTypoAscender != context.Font.HheaTable.ascender ||
                    table.sTypoDescender != context.Font.HheaTable.descender)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        "OS/2 typographic metrics differ from hhea metrics. Consider harmonizing for consistency.");
                }
            }

            // -------------------------
            // Subset-specific range check
            // -------------------------
            //if (context.SubsetChars != null && context.SubsetChars.Count > 0)
            //{
            //    int subsetMin = context.SubsetChars.Min;
            //    int subsetMax = context.SubsetChars.Max;
            //    if (subsetMin < table.usFirstCharIndex || subsetMax > table.usLastCharIndex)
            //    {
            //        result.AddMessage(FontValidationSeverity.Error,
            //            $"Subset character range ({subsetMin}-{subsetMax}) is outside OS/2 declared range ({table.usFirstCharIndex}-{table.usLastCharIndex}). Update required.");
            //    }
            //}

            return result;
        }

    }
}
