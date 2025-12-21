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
using EPPlus.Fonts.OpenType.Tables.Common.Coverage;

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

            // 1. Mappa om listorna. 
            // Vår hjälpmetod RewriteCoverageList returnerar nu en tom lista istället för null
            // om originalet var tomt/null, för att tillfredsställa Writer-koden.
            newTable.BacktrackCoverages = RewriteCoverageList(this.BacktrackCoverages, context) ?? new List<CoverageTable>();
            newTable.InputCoverages = RewriteCoverageList(this.InputCoverages, context);
            newTable.LookaheadCoverages = RewriteCoverageList(this.LookaheadCoverages, context) ?? new List<CoverageTable>();

            // 2. Validering: Input MÅSTE finnas för att regeln ska vara giltig.
            if (newTable.InputCoverages == null || newTable.InputCoverages.Count == 0) return null;

            // 3. Kopiera records
            if (this.SubstLookupRecords != null)
            {
                newTable.SubstLookupRecords = new List<SubstLookupRecord>(this.SubstLookupRecords);
            }
            else
            {
                newTable.SubstLookupRecords = new List<SubstLookupRecord>();
            }

            return newTable;
        }

        private List<CoverageTable> RewriteCoverageList(List<CoverageTable> oldCoverages, FontSubsettingContext context)
        {
            // Om originalet var null eller tomt, returnera en tom lista (inte null)
            if (oldCoverages == null || oldCoverages.Count == 0)
            {
                return new List<CoverageTable>();
            }

            var newCoverages = new List<CoverageTable>();
            foreach (var cov in oldCoverages)
            {
                var oldGids = cov.GetCoveredGlyphs();
                var validNewGids = new List<ushort>();

                foreach (var oldGid in oldGids)
                {
                    if (context.OldToNewGlyphId.TryGetValue(oldGid, out ushort newGid))
                    {
                        validNewGids.Add(newGid);
                    }
                }

                // Om en hel position i sekvensen försvinner, blir regeln ogiltig
                if (validNewGids.Count == 0) return null;

                validNewGids.Sort();
                newCoverages.Add(new CoverageTableFormat1
                {
                    GlyphArray = validNewGids.ToArray(),
                    GlyphCount = (ushort)validNewGids.Count
                });
            }
            return newCoverages;
        }
    }
}