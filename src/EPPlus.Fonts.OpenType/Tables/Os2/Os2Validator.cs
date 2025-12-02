
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
            var result = new TableValidationResult();
            result.TableName = TableName;

            // Rule 1: Version check
            if (table.version < 0 || table.version > 5)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("Unexpected OS/2 version: {0}. Expected 0–5.", table.version));
            }

            // Rule 2: usWeightClass (100–900)
            if (table.usWeightClass < 100 || table.usWeightClass > 900)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("usWeightClass out of range: {0}. Expected 100–900.", table.usWeightClass));
            }

            // Rule 3: usWidthClass (1–9)
            if (table.usWidthClass < 1 || table.usWidthClass > 9)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("usWidthClass out of range: {0}. Expected 1–9.", table.usWidthClass));
            }

            // Rule 4: fsType embedding permissions
            if ((table.fsType & 0x0002) != 0)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    "Font has restricted embedding (fsType bit 1 set).");
            }

            // Rule 5: Subscript/Superscript sizes > 0
            if (table.ySubscriptXSize <= 0 || table.ySuperscriptXSize <= 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    "Subscript/Superscript sizes should be greater than zero.");
            }

            // Rule 6: Panose check (warn if all zeros)
            bool panoseAllZero = true;
            foreach (var b in table.panose)
            {
                if (b != 0) { panoseAllZero = false; break; }
            }
            if (panoseAllZero)
            {
                result.AddMessage(FontValidationSeverity.Warning, "Panose classification is all zeros.");
            }


            // Cross-table: usFirstCharIndex/usLastCharIndex vs cmap
            if (context.Font.CmapTable != null)
            {
                int minChar = context.Font.CmapTable.GetMinCharCode();
                int maxChar = context.Font.CmapTable.GetMaxCharCode();
                if (minChar < table.usFirstCharIndex || maxChar > table.usLastCharIndex)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        string.Format("Cmap character range ({0}-{1}) is outside OS/2 declared range ({2}-{3}).",
                            minChar, maxChar, table.usFirstCharIndex, table.usLastCharIndex));
                }
            }


            // -------------------------
            // New critical rules for subsetting
            // -------------------------

            // Critical Rule A: Embedding permissions
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

            // Critical Rule B: Metrics must be valid
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


            //if (context.Font.HeadTable != null)
            //{
            //    if (table.usWinAscent < context.Font.HeadTable.Ymax || table.usWinDescent < Math.Abs(context.Font.HeadTable.Ymin))
            //    {
            //        result.AddMessage(FontValidationSeverity.Error,
            //            "usWinAscent/usWinDescent do not cover font bounding box from head table.");
            //    }
            //}


            if (context.Font.HheaTable != null)
            {
                if (table.sTypoAscender != context.Font.HheaTable.ascender ||
                    table.sTypoDescender != context.Font.HheaTable.descender)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        "OS/2 typographic metrics differ from hhea metrics. Consider harmonizing for consistency.");
                }
            }



            return result;
        }
    }
}
