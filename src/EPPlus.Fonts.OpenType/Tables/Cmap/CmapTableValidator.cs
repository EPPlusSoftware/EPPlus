using EPPlus.Fonts.OpenType.FontValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapTableValidator : TableValidatorBase<CmapTable>
    {
        public override string TableName
        {
            get { return TableNames.Cmap; }
        }

        public override Type TableType => typeof(CmapTable);

        public override TableValidationResult Validate(CmapTable cmap, FontValidationContext context)
        {

            var result = new TableValidationResult();
            result.TableName = TableName;

            // 1. Kontrollera version
            if (cmap.Version != 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "cmap version must be 0, found " + cmap.Version);
            }

            // 2. Kontrollera antal subtables

            if (cmap.NumTables != cmap.SubTables.Count)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    $"NumTables ({cmap.NumTables}) does not match unique subtable count ({cmap.SubTables.Count}). This is normal if multiple encoding records share the same subtable.");
            }



            // 3. Validera varje subtable
            foreach (var subtable in cmap.SubTables)
            {
                ValidateSubtable(subtable, result, context);
            }


            return result;
        }


        private void ValidateSubtable(CmapSubtableBase subtable, TableValidationResult result, FontValidationContext context)
        {
            switch (subtable.Format)
            {
                case 0:
                    ValidateFormat0((CmapSubtable0)subtable, result, context);
                    break;
                case 4:
                    ValidateFormat4((CmapSubtable4)subtable, result, context);
                    break;
                case 6:
                    ValidateFormat6((CmapSubtable6)subtable, result, context);
                    break;
                case 12:
                    ValidateFormat12((CmapSubtable12)subtable, result, context);
                    break;
                case 14:
                    ValidateFormat14((CmapSubtable14)subtable, result, context);
                    break;
                default:
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Unsupported cmap subtable format: {subtable.Format}");
                    break;
            }
        }


        private void ValidateFormat0(CmapSubtable0 st, TableValidationResult result, FontValidationContext context)
        {

            // 1. Check that GlyphIdArray has correct length
            if (st.GlyphIdArray.Length != 256)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Format 0: GlyphIdArray must have length 256, found {st.GlyphIdArray.Length}.");
            }

            // 2. Check that Length matches specification (header + 256 bytes)
            if (st.Length != 262)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Format 0: Length should be 262, found {st.Length}.");
            }

            // 3. Validate glyph IDs against maxp.numGlyphs
            var maxGlyphs = context.Font.MaxpTable.numGlyphs;
            for (int i = 0; i < st.GlyphIdArray.Length; i++)
            {
                byte glyphId = st.GlyphIdArray[i];
                if (glyphId >= maxGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Format 0: GlyphIdArray[{i}] = {glyphId} exceeds max glyph count ({maxGlyphs}).");
                }
            }

            // 4. Optional: Check language field for unusual values
            if (st.Language > 0xFFFF)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Format 0: Language value {st.Language} seems invalid (greater than 0xFFFF).");
            }

        }


        private void ValidateFormat4(CmapSubtable4 st, TableValidationResult result, FontValidationContext context)
        {
            int segCount = st.EndCode.Length;

            // 1. Check SegCountX2 consistency
            if (st.SegCountX2 != segCount * 2)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Format 4: SegCountX2 mismatch. Expected {segCount * 2}, found {st.SegCountX2}.");
            }

            // 2. Check that all segment arrays have equal length
            if (st.StartCode.Length != segCount || st.IdDelta.Length != segCount || st.IdRangeOffset.Length != segCount)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    "Format 4: Segment arrays must have equal length.");
            }

            // 3. Check that the last EndCode is 0xFFFF
            if (segCount == 0 || st.EndCode[segCount - 1] != 0xFFFF)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    "Format 4: Last EndCode must be 0xFFFF.");
            }

            // 4. ReservedPad should be zero
            if (st.ReservedPad != 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    "Format 4: ReservedPad should be 0.");
            }

            // 5. Validate segment ordering and ranges
            for (int i = 0; i < segCount; i++)
            {
                if (st.StartCode[i] > st.EndCode[i])
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Format 4: StartCode[{i}] ({st.StartCode[i]}) > EndCode[{i}] ({st.EndCode[i]}).");
                }

                if (i > 0 && st.StartCode[i] <= st.EndCode[i - 1])
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Format 4: Segment {i} overlaps previous segment.");
                }
            }

            // 6. Validate glyph IDs in GlyphIdArray against maxp.numGlyphs
            var maxGlyphs = context.Font.MaxpTable.numGlyphs;
            for (int i = 0; i < st.GlyphIdArray.Length; i++)
            {
                ushort glyphId = st.GlyphIdArray[i];
                if (glyphId >= maxGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Format 4: GlyphIdArray[{i}] = {glyphId} exceeds max glyph count ({maxGlyphs}).");
                }
            }

            // 7. Optional: Validate IdRangeOffset values
            for (int i = 0; i < segCount; i++)
            {
                ushort offset = st.IdRangeOffset[i];
                if (offset != 0)
                {
                    int calculatedIndex = (offset / 2) + (int)(st.StartCode[i] - st.StartCode[i]) - (segCount - i);
                    if (calculatedIndex < 0 || calculatedIndex >= st.GlyphIdArray.Length)
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            $"Format 4: IdRangeOffset[{i}] points outside GlyphIdArray.");
                    }
                }
            }


            // 8. Validate searchRange, entrySelector, rangeShift
            int maxPower = (int)Math.Pow(2, (int)Math.Floor(Math.Log(segCount, 2)));
            ushort expectedSearchRange = (ushort)(maxPower * 2);
            ushort expectedEntrySelector = (ushort)Math.Floor(Math.Log(segCount, 2));
            ushort expectedRangeShift = (ushort)((segCount * 2) - expectedSearchRange);

            if (st.SearchRange != expectedSearchRange)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Format 4: searchRange mismatch. Expected {expectedSearchRange}, found {st.SearchRange}.");
            }
            if (st.EntrySelector != expectedEntrySelector)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Format 4: entrySelector mismatch. Expected {expectedEntrySelector}, found {st.EntrySelector}.");
            }
            if (st.RangeShift != expectedRangeShift)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Format 4: rangeShift mismatch. Expected {expectedRangeShift}, found {st.RangeShift}.");
            }

        }




        private void ValidateFormat6(CmapSubtable6 st, TableValidationResult result, FontValidationContext context)
        {
            // 1. Check that GlyphIdArray length matches EntryCount
            if (st.GlyphIdArray.Length != st.EntryCount)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Format 6: GlyphIdArray length ({st.GlyphIdArray.Length}) does not match EntryCount ({st.EntryCount}).");
            }

            // 2. Check that Length matches specification: header (10 bytes) + glyph array (EntryCount * 2)
            uint expectedLength = 10 + (uint)st.EntryCount * 2;
            if (st.Length != expectedLength)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Format 6: Length should be {expectedLength}, found {st.Length}.");
            }

            // 3. Check that the character range does not exceed 0xFFFF
            uint lastCode = (uint)st.FirstCode + (uint)st.EntryCount - 1;
            if (lastCode > 0xFFFF)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Format 6: Character range exceeds 0xFFFF. FirstCode={st.FirstCode}, EntryCount={st.EntryCount}.");
            }

            // 4. Validate that all glyph IDs are within the font's glyph count
            var maxGlyphs = context.Font.MaxpTable.numGlyphs;
            for (int i = 0; i < st.GlyphIdArray.Length; i++)
            {
                ushort glyphId = st.GlyphIdArray[i];
                if (glyphId >= maxGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Format 6: GlyphIdArray[{i}] = {glyphId} exceeds max glyph count ({maxGlyphs}).");
                }
            }
        }



        private void ValidateFormat12(CmapSubtable12 st, TableValidationResult result, FontValidationContext context)
        {
            // 1. Check that NumGroups matches actual group count
            if (st.Groups.Count != st.NumGroups)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Format 12: NumGroups ({st.NumGroups}) does not match actual group count ({st.Groups.Count}).");
            }

            // 2. Validate each group for logical consistency
            for (int i = 0; i < st.Groups.Count; i++)
            {
                var g = st.Groups[i];

                // StartCharCode must not be greater than EndCharCode
                if (g.StartCharCode > g.EndCharCode)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Format 12: Group {i} StartCharCode ({g.StartCharCode}) > EndCharCode ({g.EndCharCode}).");
                }

                // Groups must not overlap previous group
                if (i > 0 && g.StartCharCode <= st.Groups[i - 1].EndCharCode)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Format 12: Group {i} overlaps previous group. StartCharCode={g.StartCharCode}, PreviousEnd={st.Groups[i - 1].EndCharCode}.");
                }

                // Validate glyph range against maxp.numGlyphs
                var maxGlyphs = context.Font.MaxpTable.numGlyphs;
                ulong glyphRangeEnd = (ulong)g.StartGlyphId + (ulong)(g.EndCharCode - g.StartCharCode);
                if (glyphRangeEnd >= (ulong)maxGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Format 12: Group {i} glyph range exceeds max glyph count ({maxGlyphs}). StartGlyphId={g.StartGlyphId}, EndGlyphId={glyphRangeEnd}.");
                }

                // Optional: Validate that character codes do not exceed Unicode max (0x10FFFF)
                if (g.EndCharCode > 0x10FFFF)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Format 12: Group {i} EndCharCode ({g.EndCharCode}) exceeds Unicode maximum (0x10FFFF).");
                }
            }
        }





        private void ValidateFormat14(CmapSubtable14 st, TableValidationResult result, FontValidationContext context)
        {
            // 1. Check that VariationSelectors are sorted
            for (int i = 1; i < st.VariationSelectors.Count; i++)
            {
                if (st.VariationSelectors[i].VarSelector < st.VariationSelectors[i - 1].VarSelector)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Format 14: VariationSelectors not sorted at index {i}. VarSelector={st.VariationSelectors[i].VarSelector:X6}, previous={st.VariationSelectors[i - 1].VarSelector:X6}.");
                }
            }

            foreach (var selector in st.VariationSelectors)
            {
                uint vs = selector.VarSelector;

                // 2. Validate VarSelector range
                bool isValidRange = (vs >= 0xFE00 && vs <= 0xFE0F) || (vs >= 0xE0100 && vs <= 0xE01EF);
                if (!isValidRange)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Format 14: VarSelector {vs:X6} is outside standard Unicode variation selector ranges.");
                }

                // 3. Validate NonDefault UVS mappings
                if (selector.NonDefaultUvsTable != null)
                {
                    var seenUnicodeValues = new HashSet<uint>();
                    foreach (var entry in selector.NonDefaultUvsTable.Mappings)
                    {
                        if (!seenUnicodeValues.Add(entry.UnicodeValue))
                        {
                            result.AddMessage(FontValidationSeverity.Error,
                                $"Format 14: Duplicate UnicodeValue {entry.UnicodeValue:X6} in NonDefault UVS table for VarSelector {vs:X6}.");
                        }

                        var maxGlyphs = context.Font.MaxpTable.numGlyphs;
                        if (entry.GlyphId >= maxGlyphs)
                        {
                            result.AddMessage(FontValidationSeverity.Error,
                                $"Format 14: GlyphId {entry.GlyphId} exceeds max glyph count ({maxGlyphs}) for UnicodeValue {entry.UnicodeValue:X6}.");
                        }
                    }
                }

                // 4. Validate Default UVS ranges
                if (selector.DefaultUvsTable != null)
                {
                    foreach (var range in selector.DefaultUvsTable.Ranges)
                    {
                        uint start = range.StartUnicodeValue;
                        uint end = start + range.AdditionalCount;

                        // Check start and end against Unicode max
                        if (start > 0x10FFFF)
                        {
                            result.AddMessage(FontValidationSeverity.Warning,
                                $"Format 14: Default UVS range start {start:X6} exceeds Unicode maximum (0x10FFFF).");
                        }
                        if (end > 0x10FFFF)
                        {
                            result.AddMessage(FontValidationSeverity.Warning,
                                $"Format 14: Default UVS range end {end:X6} exceeds Unicode maximum (0x10FFFF).");
                        }

                        // Optional: Check that ranges are sorted and non-overlapping
                        // (We can implement this if needed for stricter validation)
                    }
                }
            }
        }
    }
}
