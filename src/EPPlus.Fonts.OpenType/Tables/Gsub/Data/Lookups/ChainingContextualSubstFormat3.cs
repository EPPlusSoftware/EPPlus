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
using EPPlus.Fonts.OpenType.Subsetting;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    public class ChainingContextualSubstFormat3 : FontTableElement
    {
        public List<CoverageTable> BacktrackCoverages { get; set; } = new();
        public List<CoverageTable> InputCoverages { get; set; } = new();
        public List<CoverageTable> LookaheadCoverages { get; set; } = new();
        public List<SubstLookupRecord> SubstLookupRecords { get; set; } = new();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long startPos = writer.BaseStream.Position;

            // 1. Format 3
            writer.WriteUInt16BigEndian(3);

            // 2. Backtrack Coverage
            writer.WriteUInt16BigEndian((ushort)BacktrackCoverages.Count);
            long backtrackOffsetsStart = writer.BaseStream.Position;
            for (int i = 0; i < BacktrackCoverages.Count; i++) writer.WriteUInt16BigEndian(0);

            // 3. Input Coverage
            writer.WriteUInt16BigEndian((ushort)InputCoverages.Count);
            long inputOffsetsStart = writer.BaseStream.Position;
            for (int i = 0; i < InputCoverages.Count; i++) writer.WriteUInt16BigEndian(0);

            // 4. Lookahead Coverage
            writer.WriteUInt16BigEndian((ushort)LookaheadCoverages.Count);
            long lookaheadOffsetsStart = writer.BaseStream.Position;
            for (int i = 0; i < LookaheadCoverages.Count; i++) writer.WriteUInt16BigEndian(0);

            // 5. SubstLookupRecords
            writer.WriteUInt16BigEndian((ushort)SubstLookupRecords.Count);
            foreach (var record in SubstLookupRecords)
            {
                writer.WriteUInt16BigEndian(record.SequenceIndex);
                writer.WriteUInt16BigEndian(record.LookupListIndex);
            }

            // --- Skriv ut faktiska Coverage-tabeller och backfilla offsets ---

            SerializeCoverageList(writer, BacktrackCoverages, startPos, backtrackOffsetsStart);
            SerializeCoverageList(writer, InputCoverages, startPos, inputOffsetsStart);
            SerializeCoverageList(writer, LookaheadCoverages, startPos, lookaheadOffsetsStart);
        }

        private void SerializeCoverageList(FontsBinaryWriter writer, List<CoverageTable> coverages, long subTableStart, long offsetArrayStart)
        {
            for (int i = 0; i < coverages.Count; i++)
            {
                long currentPos = writer.BaseStream.Position;
                long offsetInArray = offsetArrayStart + (i * 2);

                this.WriteRelativeOffset(writer, subTableStart, offsetInArray);
                coverages[i].Serialize(writer);
            }
        }

        internal ChainingContextualSubstFormat3 Rewrite(FontSubsettingContext context)
        {
            var newTable = new ChainingContextualSubstFormat3();

            // För varje kontext-del (Backtrack, Input, Lookahead) måste vi kontrollera 
            // om alla tecken i Coverage fortfarande är relevanta.

            newTable.BacktrackCoverages = RewriteCoverageList(BacktrackCoverages, context);
            newTable.InputCoverages = RewriteCoverageList(InputCoverages, context);
            newTable.LookaheadCoverages = RewriteCoverageList(LookaheadCoverages, context);

            // Om någon av listorna blev tom (men inte var det från början), 
            // eller om den kritiska Input-listan är bruten, kasta hela subtabellen.
            if (newTable.InputCoverages == null || newTable.InputCoverages.Count == 0) return null;
            if (BacktrackCoverages.Count > 0 && (newTable.BacktrackCoverages == null || newTable.BacktrackCoverages.Count == 0)) return null;
            if (LookaheadCoverages.Count > 0 && (newTable.LookaheadCoverages == null || newTable.LookaheadCoverages.Count == 0)) return null;

            // Behåll lookup-records som de är (de pekar på LookupList-index, 
            // vilket hanteras i LookupListTable.Rewrite)
            newTable.SubstLookupRecords = new List<SubstLookupRecord>(this.SubstLookupRecords);

            return newTable;
        }

        private List<CoverageTable> RewriteCoverageList(List<CoverageTable> oldCoverages, FontSubsettingContext context)
        {
            var newCoverages = new List<CoverageTable>();
            foreach (var cov in oldCoverages)
            {
                // Skapa en ny CoverageTable som bara innehåller de GIDs som finns i vårt subset
                var filteredGids = cov.GetCoveredGlyphs()
                    .Where(oldGid => context.OldToNewGlyphId.ContainsKey(oldGid))
                    .Select(oldGid => context.OldToNewGlyphId[oldGid])
                    .OrderBy(newGid => newGid)
                    .ToList();

                if (filteredGids.Count == 0) return null; // En del av kontexten försvann helt!

                newCoverages.Add(new CoverageTableFormat1
                {
                    GlyphCount = (ushort)filteredGids.Count,
                    GlyphArray = filteredGids.ToArray()
                });
            }
            return newCoverages;
        }
    }
}