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
        public void Discover(FontSubsettingContext context)
        {
            // 1. Always include .notdef (GID 0)
            context.IncludedGlyphs.Add(0);

            // 2. Recursively find and include all components for composite glyphs
            // This is critical for fonts like Times New Roman
            context.OriginalFont.GlyfTable.ResolveCompositeGlyphs(context.IncludedGlyphs);
        }

        public void Rewrite(FontSubsettingContext context)
        {
            // NewToOldGlyphId is sorted by new IDs (0, 1, 2...)
            var sortedOldIds = context.NewToOldGlyphId;
            List<Glyph> newGlyphs = new List<Glyph>(sortedOldIds.Count);

            // 1. Clone and remap glyphs
            foreach (ushort oldId in sortedOldIds)
            {
                Glyph originalGlyph = context.OriginalFont.GlyfTable.GetGlyph(oldId);

                if (originalGlyph == null || IsEmpty(originalGlyph))
                {
                    newGlyphs.Add(new Glyph { Header = new GlyphHeader(0, BoundingRectangle.Empty) });
                }
                else
                {
                    newGlyphs.Add(CloneGlyphWithRemappedComponents(originalGlyph, context.OldToNewGlyphId));
                }
            }

            // 2. Save the new glyf table
            context.SubsetFont.AddOrReplaceTable(new GlyfTable(newGlyphs));

            // 3. Build loca table with 4-byte alignment
            List<uint> offsets = new List<uint> { 0 };
            uint currentOffset = 0;

            foreach (Glyph g in newGlyphs)
            {
                int size = g.GetSize();
                uint paddedSize = (uint)((size + 3) & ~3); // Align to 4 bytes
                currentOffset += paddedSize;
                offsets.Add(currentOffset);
            }

            // 4. Update head table format and add loca table
            bool useShortOffsets = currentOffset <= 131070;
            context.SubsetFont.HeadTable.IndexToLocFormat = useShortOffsets
                ? HeadTable.IndexToLocFormats.Offset16
                : HeadTable.IndexToLocFormats.Offset32;

            context.SubsetFont.AddOrReplaceTable(LocaTable.CreateSubset(offsets, context.SubsetFont.HeadTable.IndexToLocFormat));
        }

        private static bool IsEmpty(Glyph g)
        {
            return g.Header.numberOfContours == 0 && g.CompositeData == null && g.SimpleData == null;
        }

        private static Glyph CloneGlyphWithRemappedComponents(Glyph original, Dictionary<ushort, ushort> oldToNewMap)
        {
            GlyphHeader header = new GlyphHeader(original.Header.numberOfContours, original.Header.Bounds);

            // Handle Simple Glyph
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

            // Handle Composite Glyph
            if (original.Header.numberOfContours < 0 && original.CompositeData != null)
            {
                CompositeGlyph composite = new CompositeGlyph
                {
                    Instructions = (byte[])original.CompositeData.Instructions.Clone(),
                    Components = new List<GlyphComponent>()
                };

                foreach (GlyphComponent comp in original.CompositeData.Components)
                {
                    // Remap the component's GlyphIndex to the new ID
                    if (!oldToNewMap.TryGetValue(comp.GlyphIndex, out ushort newGid))
                    {
                        newGid = 0; // Fallback to .notdef
                    }

                    composite.Components.Add(new GlyphComponent
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
                    });
                }

                return new Glyph { Header = header, CompositeData = composite };
            }

            return new Glyph { Header = header };
        }
    }
}
