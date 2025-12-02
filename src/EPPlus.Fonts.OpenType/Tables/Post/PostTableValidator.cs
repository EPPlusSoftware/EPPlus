using EPPlus.Fonts.OpenType.FontValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Fonts.OpenType.Tables.Post
{
    // All comments are in English
    internal class PostTableValidator : TableValidatorBase<PostTable>
    {
        public override string TableName { get { return TableNames.Post; } }
        public override Type TableType { get { return typeof(PostTable); } }

        public override TableValidationResult Validate(PostTable table, FontValidationContext context)
        {
            TableValidationResult result = new TableValidationResult();
            result.TableName = TableName;

            // ---------- Rule 1: Version check ----------

            int rawVersion = table.version.RawValue; // int, not uint
            bool isV10 = rawVersion == PostTableConstants.Version10;
            bool isV20 = rawVersion == PostTableConstants.Version20;
            bool isV25 = rawVersion == PostTableConstants.Version25;
            bool isV30 = rawVersion == PostTableConstants.Version30;

            if (!isV10 && !isV20 && !isV25 && !isV30)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("Invalid post version: 0x{0:X}. Expected 0x00010000 (1.0), 0x00020000 (2.0), 0x00025000 (2.5) or 0x00030000 (3.0).", rawVersion));
            }


            // ---------- Rule 2: isFixedPitch must be 0 or 1 ----------
            if (table.isFixedPitch != 0 && table.isFixedPitch != 1)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("isFixedPitch is {0}. Expected 0 (proportional) or 1 (monospaced).", table.isFixedPitch));
            }

            // ---------- Rule 3: Underline thickness ----------
            if (table.underlineThickness <= 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("underlineThickness ({0}) should be greater than 0.", table.underlineThickness));
            }
            else
            {
                // Optional: compare to unitsPerEm (soft check)
                if (context != null && context.Font != null && context.Font.HeadTable != null)
                {
                    int unitsPerEm = context.Font.HeadTable.UnitsPerEm;
                    if (table.underlineThickness > unitsPerEm)
                    {
                        result.AddMessage(FontValidationSeverity.Information,
                            string.Format("underlineThickness ({0}) exceeds unitsPerEm ({1}).", table.underlineThickness, unitsPerEm));
                    }
                }
            }

            // ---------- Rule 4: Italic angle sanity (soft) ----------
            // Fixed16Dot16 likely exposes double or float; assume .ToDouble() or similar. If not, adjust accordingly.
            float italicAngleValue = table.italicAngle.FloatValue;
            if (Math.Abs(italicAngleValue) > PostTableConstants.ItalicAngleMaxAbsDegreesWarning)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("italicAngle ({0}) is unusually large.", italicAngleValue));
            }

            // ---------- Rule 5: Version-specific checks ----------
            if (isV20)
            {
                // v2.0 has glyph names via glyphNameIndex and Pascal strings (for indices >= 258)
                // Critical: numGlyphs must be > 0 and align with name arrays
                if (table.numGlyphs <= 0)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        string.Format("post v2.0: numGlyphs is {0}. Must be greater than 0.", table.numGlyphs));
                }

                // Arrays must be present and aligned (your serializer accesses glyphNames[i] for all i)
                if (table.glyphNameIndex == null || table.glyphNameIndex.Count != table.numGlyphs)
                {
                    int actual = (table.glyphNameIndex == null) ? 0 : table.glyphNameIndex.Count;
                    result.AddMessage(FontValidationSeverity.Error,
                        string.Format("post v2.0: glyphNameIndex count ({0}) does not match numGlyphs ({1}).", actual, table.numGlyphs));
                }

                if (table.glyphNames == null || table.glyphNames.Count != table.numGlyphs)
                {
                    int actual = (table.glyphNames == null) ? 0 : table.glyphNames.Count;
                    result.AddMessage(FontValidationSeverity.Error,
                        string.Format("post v2.0: glyphNames count ({0}) does not match numGlyphs ({1}).", actual, table.numGlyphs));
                }

                // Cross-table critical: must match maxp.numGlyphs
                if (context != null && context.Font != null && context.Font.MaxpTable != null)
                {
                    if (context.Font.MaxpTable.numGlyphs != table.numGlyphs)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            string.Format("post v2.0: numGlyphs ({0}) does not match maxp.numGlyphs ({1}).",
                                table.numGlyphs, context.Font.MaxpTable.numGlyphs));
                    }
                }

                // Validate per-entry rules
                if (table.glyphNameIndex != null && table.glyphNames != null &&
                    table.glyphNameIndex.Count == table.glyphNames.Count)
                {
                    // Count custom names (index >= 258)
                    int customCount = 0;

                    for (int i = 0; i < table.glyphNameIndex.Count; i++)
                    {
                        ushort idx = table.glyphNameIndex[i];
                        string name = table.glyphNames[i];

                        if (idx >= PostTableConstants.StandardMacGlyphNameCount)
                        {
                            customCount++;

                            // Custom names must be ASCII and <= 255 bytes (Pascal length limit)
                            if (name == null)
                            {
                                result.AddMessage(FontValidationSeverity.Error,
                                    string.Format("post v2.0: glyphNames[{0}] is null for custom index {1}.", i, idx));
                                continue;
                            }

                            byte[] asciiBytes = System.Text.Encoding.ASCII.GetBytes(name);
                            if (asciiBytes.Length > 255)
                            {
                                result.AddMessage(FontValidationSeverity.Error,
                                    string.Format("post v2.0: glyphNames[{0}] ('{1}') exceeds 255 bytes.", i, name));
                            }

                            // Soft check: ensure all characters are 7-bit ASCII
                            for (int b = 0; b < asciiBytes.Length; b++)
                            {
                                if (asciiBytes[b] > 0x7F)
                                {
                                    result.AddMessage(FontValidationSeverity.Warning,
                                        string.Format("post v2.0: glyphNames[{0}] ('{1}') contains non-ASCII characters.", i, name));
                                    break;
                                }
                            }
                        }
                        else
                        {
                            // idx < 258 uses standard Mac glyph name list; custom string should not be serialized.
                            // Your serializer still reads glyphNames[i] but only writes for idx >= 258.
                            // If you want stricter hygiene, warn when a standard index has a non-empty custom name.
                            if (!string.IsNullOrEmpty(name))
                            {
                                result.AddMessage(FontValidationSeverity.Information,
                                    string.Format("post v2.0: glyphNames[{0}] is non-empty ('{1}') for standard index {2}. It will not be serialized.", i, name, idx));
                            }
                        }
                    }

                    // Optional sanity: the number of custom indices should match the count of non-empty custom names.
                    // Since you store a glyphNames entry for each glyph, we used presence of idx >= 258 as the authoritative rule.
                }
            }
            else if (isV25)
            {
                // v2.5: glyph name indices replaced with offsets to standard Mac names; no custom Pascal strings.
                // Given your data model does not expose specific 2.5 arrays, we at least warn if arrays are populated.
                if ((table.glyphNameIndex != null && table.glyphNameIndex.Count > 0) ||
                    (table.glyphNames != null && table.glyphNames.Count > 0) ||
                    table.numGlyphs > 0)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        "post v2.5: Custom glyph name arrays are not expected; data will be ignored.");
                }
            }
            else if (isV30)
            {
                // v3.0: no glyph names are stored. Names are not provided; recommend arrays empty.
                if ((table.glyphNameIndex != null && table.glyphNameIndex.Count > 0) ||
                    (table.glyphNames != null && table.glyphNames.Count > 0) ||
                    table.numGlyphs > 0)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        "post v3.0: Glyph names should not be present; arrays should be empty.");
                }
            }
            else if (isV10)
            {
                // v1.0: legacy; no glyph names arrays defined. Keep soft check similar to v3.0.
                if ((table.glyphNameIndex != null && table.glyphNameIndex.Count > 0) ||
                    (table.glyphNames != null && table.glyphNames.Count > 0) ||
                    table.numGlyphs > 0)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        "post v1.0: Glyph name arrays are not expected; arrays should be empty.");
                }
            }

            // ---------- Rule 6: Memory fields consistency ----------
            // In many modern fonts, these are zeroed—especially in v3.0. We keep this as a soft check.
            if (isV30)
            {
                if (table.minMemType42 != 0 || table.maxMemType42 != 0 ||
                    table.minMemType1 != 0 || table.maxMemType1 != 0)
                {
                    result.AddMessage(FontValidationSeverity.Information,
                        "post v3.0: min/max memory fields are typically zero in modern fonts.");
                }
            }

            return result;
        }
    }
}

