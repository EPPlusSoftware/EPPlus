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
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Loca;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Subsets the glyf and loca tables, builds glyph ID remapping, and resolves composite glyphs.
    /// This processor must run early in the pipeline (after code point collection).
    /// Fully compatible with .NET 3.5.
    /// </summary>
    internal class GlyfAndLocaSubsetProcessor : IFontSubsetProcessor
    {
        public void Process(FontSubsettingContext context)
        {
            OpenTypeFont originalFont = context.OriginalFont;
            OpenTypeFont subsetFont = context.SubsetFont;

            // --------------------------------------------------------------------
            // 1. Build initial glyph set from used code points
            // --------------------------------------------------------------------
            foreach (uint codePoint in context.UsedCodePoints)
            {
                if (originalFont.CmapTable.TryGetGlyphId(codePoint, out ushort gid))
                {
                    context.IncludedGlyphs.Add(gid);
                }
            }

            // --------------------------------------------------------------------
            // 2. Ensure .notdef (GID 0) is always included
            // --------------------------------------------------------------------
            context.IncludedGlyphs.Add(0);

            // --------------------------------------------------------------------
            // 3. Resolve all composite glyph components recursively
            // --------------------------------------------------------------------
            originalFont.GlyfTable.ResolveCompositeGlyphs(context.IncludedGlyphs);

            // --------------------------------------------------------------------
            // 4. Sort glyph IDs and ensure .notdef is exactly once at position 0
            // --------------------------------------------------------------------
            List<ushort> sortedGlyphIds = new List<ushort>(context.IncludedGlyphs);
            sortedGlyphIds.Sort();

            // Remove any existing .notdef entries (in case it was added multiple times)
            sortedGlyphIds.RemoveAll(g => g == 0);

            // Insert .notdef exactly once as the first glyph (required by spec)
            sortedGlyphIds.Insert(0, 0);

            // --------------------------------------------------------------------
            // 5. Build old to new and new to old glyph ID mappings
            // --------------------------------------------------------------------
            context.OldToNewGlyphId.Clear();
            context.NewToOldGlyphId.Clear();

            for (int newId = 0; newId < sortedGlyphIds.Count; newId++)
            {
                ushort oldId = sortedGlyphIds[newId];
                context.OldToNewGlyphId[oldId] = (ushort)newId;
                context.NewToOldGlyphId.Add(oldId);
            }

            // --------------------------------------------------------------------
            // 6. Clone glyphs and remap component references (for composites)
            // --------------------------------------------------------------------
            List<Glyph> newGlyphs = new List<Glyph>(sortedGlyphIds.Count);

            foreach (ushort oldId in sortedGlyphIds)
            {
                Glyph originalGlyph = originalFont.GlyfTable.GetGlyph(oldId);

                if (originalGlyph == null || originalGlyph.Header.numberOfContours == 0)
                {
                    // Empty glyph (e.g. space, .null) – only header is needed
                    newGlyphs.Add(new Glyph
                    {
                        Header = new GlyphHeader(0, BoundingRectangle.Empty)
                    });
                }
                else
                {
                    newGlyphs.Add(CloneGlyphWithRemappedComponents(originalGlyph, context.OldToNewGlyphId));
                }
            }

            subsetFont.AddOrReplaceTable(new GlyfTable(newGlyphs));

            // --------------------------------------------------------------------
            // 7. Build loca table with 4-byte alignment (short or long format)
            // --------------------------------------------------------------------
            List<uint> offsets = new List<uint> { 0 };
            uint currentOffset = 0;

            foreach (Glyph g in newGlyphs)
            {
                int size = g.GetSize();
                currentOffset += (uint)((size + 3) & ~3); // 4-byte padding
                offsets.Add(currentOffset);
            }

            // Use short loca if all offsets fit in 16 bits
            bool useShortOffsets = true;
            foreach (uint offset in offsets)
            {
                if (offset > 131070)
                {
                    useShortOffsets = false;
                    break;
                }
            }

            subsetFont.HeadTable.IndexToLocFormat = useShortOffsets
                ? HeadTable.IndexToLocFormats.Offset16
                : HeadTable.IndexToLocFormats.Offset32;

            subsetFont.AddOrReplaceTable(LocaTable.CreateSubset(offsets, subsetFont.HeadTable.IndexToLocFormat));
        }

        /// <summary>
        /// Clones a glyph and remaps all composite component glyph IDs.
        /// Preserves bounding box, instructions, flags and transformation matrices.
        /// .NET 3.5 compatible – uses only Dictionary&lt;ushort, ushort&gt;.
        /// </summary>
        private static Glyph CloneGlyphWithRemappedComponents(Glyph original, Dictionary<ushort, ushort> oldToNewMap)
        {
            // Fix inverted bounding boxes (common in buggy .notdef glyphs)
            short xMin = original.Header.xMin;
            short xMax = original.Header.xMax;
            short yMin = original.Header.yMin;
            short yMax = original.Header.yMax;

            if (xMin > xMax) { short tmp = xMin; xMin = xMax; xMax = tmp; }
            if (yMin > yMax) { short tmp = yMin; yMin = yMax; yMax = tmp; }

            GlyphHeader header = new GlyphHeader(original.Header.numberOfContours,
                                                new BoundingRectangle(xMin, yMin, xMax, yMax));

            // Simple (contour-based) glyph
            if (original.Header.numberOfContours > 0 && original.SimpleData != null)
            {
                SimpleGlyph simple = new SimpleGlyph
                {
                    EndPtsOfContours = (ushort[])original.SimpleData.EndPtsOfContours.Clone(),
                    Instructions = (byte[])original.SimpleData.Instructions.Clone(),
                    XBytes = (byte[])original.SimpleData.XBytes.Clone(),
                    YBytes = (byte[])original.SimpleData.YBytes.Clone(),
                    FlagRuns = new List<FlagRun>()
                };

                foreach (FlagRun run in original.SimpleData.FlagRuns)
                {
                    simple.FlagRuns.Add(new FlagRun { Flag = run.Flag, RepeatCount = run.RepeatCount });
                }

                return new Glyph { Header = header, SimpleData = simple };
            }

            // Composite glyph – remap all components
            if (original.Header.numberOfContours < 0 && original.CompositeData != null)
            {
                CompositeGlyph composite = new CompositeGlyph
                {
                    Instructions = (byte[])original.CompositeData.Instructions.Clone(),
                    Components = new List<GlyphComponent>()
                };

                foreach (GlyphComponent comp in original.CompositeData.Components)
                {
                    ushort newGid = 0;
                    if (!oldToNewMap.TryGetValue(comp.GlyphIndex, out newGid))
                    {
                        // Should never happen – fallback to .notdef
                        newGid = 0;
                    }

                    GlyphComponent newComp = new GlyphComponent
                    {
                        Flags = comp.Flags,
                        GlyphIndex = newGid,
                        Argument1 = comp.Argument1,
                        Argument2 = comp.Argument2,
                        Scale = comp.Scale,
                        XScale = comp.XScale,
                        YScale = comp.YScale,
                        Scale01 = comp.Scale01,
                        Scale10 = comp.Scale10
                    };

                    composite.Components.Add(newComp);
                }

                return new Glyph { Header = header, CompositeData = composite };
            }

            // Empty glyph (e.g. space)
            return new Glyph { Header = header };
        }
    }
}
