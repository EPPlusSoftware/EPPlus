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
using System.Collections.Generic;
using EPPlus.Fonts.OpenType.Subsetting;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups
{
    /// <summary>
    /// Represents the Lookup List table in GSUB, which contains all the lookups used for substitutions.
    /// </summary>
    public class LookupListTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the list of lookups.
        /// </summary>
        public List<LookupTable> Lookups { get; set; } = new List<LookupTable>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long startPos = writer.BaseStream.Position;

            // 1. Write LookupCount
            writer.WriteUInt16BigEndian((ushort)Lookups.Count);

            // 2. Placeholders for LookupOffsets
            long offsetArrayStart = writer.BaseStream.Position;
            for (int i = 0; i < Lookups.Count; i++)
            {
                writer.WriteUInt16BigEndian(0);
            }

            // 3. Serialize Lookups and backfill offsets
            for (int i = 0; i < Lookups.Count; i++)
            {
                long currentPos = writer.BaseStream.Position;
                long offsetInArray = offsetArrayStart + (i * 2);

                this.WriteRelativeOffset(writer, startPos, offsetInArray);
                Lookups[i].Serialize(writer);
            }
        }

        /// <summary>
        /// Rewrites the lookup list. Note that in a full implementation, 
        /// removing lookups might require remapping indexes in Features.
        /// </summary>
        internal LookupRewriteResult Rewrite(FontSubsettingContext context)
        {
            var result = new LookupRewriteResult
            {
                NewLookupList = new LookupListTable(),
                OldToNewIndexMap = new Dictionary<int, int>()
            };

            for (int i = 0; i < this.Lookups.Count; i++)
            {
                var oldLookup = this.Lookups[i];

                var rewrittenLookup = context.GsubProcessor.RewriteLookup(context, oldLookup);

                if (rewrittenLookup != null && rewrittenLookup.SubTables.Count > 0)
                {
                    int newIndex = result.NewLookupList.Lookups.Count;
                    result.NewLookupList.Lookups.Add(rewrittenLookup);
                    result.OldToNewIndexMap[i] = newIndex;
                }
            }

            // Contextual subtables still reference lookups by their original index. The new
            // lookup order is only known now, after all lookups have been rewritten.
            result.NewLookupList.RemapLookupRecords(result.OldToNewIndexMap);

            return result;
        }

        /// <summary>
        /// Remaps lookup records in all contextual subtables of this list, including subtables
        /// wrapped in Extension lookups. Must only be called on a rewritten (subset) list.
        /// </summary>
        internal void RemapLookupRecords(Dictionary<int, int> oldToNewIndexMap)
        {
            foreach (LookupTable lookup in this.Lookups)
            {
                foreach (FontTableElement subTable in lookup.SubTables)
                {
                    FontTableElement target = subTable;
                    ExtensionSubTableBase extension = target as ExtensionSubTableBase;
                    if (extension != null)
                    {
                        target = extension.ExtendedSubTable;
                    }

                    ILookupRecordContainer container = target as ILookupRecordContainer;
                    if (container != null)
                    {
                        container.RemapLookupListIndices(oldToNewIndexMap);
                    }
                }
            }
        }
    }
}