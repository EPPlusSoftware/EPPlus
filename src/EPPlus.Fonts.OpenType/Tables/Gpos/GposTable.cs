/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           GPOS table implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Features;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Utils;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gpos
{   
    
    /// <summary>
    /// Represents the GPOS (Glyph Positioning) table.
    /// Used for advanced typographic positioning including kerning, mark placement, etc.
    /// </summary>
    public class GposTable : FontTableBase
    {
        /// <summary>
        /// Major version (should be 1)
        /// </summary>
        public ushort MajorVersion { get; set; }

        /// <summary>
        /// Minor version (0 or 1)
        /// </summary>
        public ushort MinorVersion { get; set; }

        /// <summary>
        /// Script list containing language systems
        /// </summary>
        public ScriptListTable ScriptList { get; set; }

        /// <summary>
        /// Feature list containing typographic features (kern, mark, mkmk, etc)
        /// </summary>
        public FeatureListTable FeatureList { get; set; }

        /// <summary>
        /// Lookup list containing positioning rules
        /// </summary>
        public LookupListTable LookupList { get; set; }

        /// <summary>
        /// Feature variations table (GPOS 1.1)
        /// </summary>
        public uint FeatureVariationsOffset { get; set; }

        public override string Name => TableNames.Gpos;

        public override bool IsEssentialTable => false;

        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            long tableStartOffset = writer.BaseStream.Position;

            // --- HEADER ---
            writer.WriteUInt16BigEndian(this.MajorVersion);
            writer.WriteUInt16BigEndian(this.MinorVersion);

            long scriptListOffPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Placeholder

            long featureListOffPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Placeholder

            long lookupListOffPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Placeholder

            // GPOS 1.1 FeatureVariations (not implemented yet)
            if (this.MajorVersion == 1 && this.MinorVersion == 1)
            {
                writer.WriteUInt32BigEndian(0); // FeatureVariations offset (0 = not present)
            }

            // --- DATA ---

            // 1. ScriptList
            if (this.ScriptList != null)
            {
                LayoutTableSerializationHelper.UpdateOffsetAndSerialize(writer, tableStartOffset, scriptListOffPos, this.ScriptList);
            }

            // 2. FeatureList
            if (this.FeatureList != null)
            {
                LayoutTableSerializationHelper.UpdateOffsetAndSerialize(writer, tableStartOffset, featureListOffPos, this.FeatureList);
            }

            // 3. LookupList
            if (this.LookupList != null)
            {
                LayoutTableSerializationHelper.UpdateOffsetAndSerialize(writer, tableStartOffset, lookupListOffPos, this.LookupList);
            }
        }

        internal override void Clear()
        {
            // TODO: Implement when needed for memory management
            // This would clear internal caches/state if any
            ScriptList = null;
            FeatureList = null;
            LookupList = null;
        }

        /// Rewrites the GPOS table for subsetting.
        /// Filters lookups, features, and scripts based on included glyphs.
        /// </summary>
        /// <param name="context">Subsetting context</param>
        /// <returns>Rewritten GPOS table, or null if no positioning data remains</returns>
        internal GposTable Rewrite(FontSubsettingContext context)
        {
            var processor = context.GposProcessor;
            if (processor == null)
                return null;

            // Rewrite LookupList
            var newLookupList = RewriteLookupList(context, processor);
            if (newLookupList == null || newLookupList.Lookups.Count == 0)
                return null; // No positioning data remains

            // Create new GPOS table
            var newGpos = new GposTable
            {
                MajorVersion = this.MajorVersion,
                MinorVersion = this.MinorVersion,
                ScriptList = this.ScriptList,    // Keep all scripts
                FeatureList = this.FeatureList,  // Keep all features
                LookupList = newLookupList
            };

            return newGpos;
        }

        /// <summary>
        /// Rewrites the LookupList by processing each lookup through its handler.
        /// </summary>
        private LookupListTable RewriteLookupList(FontSubsettingContext context, GposSubsetProcessor processor)
        {
            if (this.LookupList == null)
                return null;

            var newLookups = new List<LookupTable>();

            foreach (var lookup in this.LookupList.Lookups)
            {
                var rewrittenLookup = processor.RewriteLookup(context, lookup);
                if (rewrittenLookup != null && rewrittenLookup.SubTables.Count > 0)
                {
                    newLookups.Add(rewrittenLookup);
                }
            }

            if (newLookups.Count == 0)
                return null;

            return new LookupListTable
            {
                Lookups = newLookups
            };
        }
    }
}
