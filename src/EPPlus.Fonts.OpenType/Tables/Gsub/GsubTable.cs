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
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Features;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Utils;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class GsubTable : FontTableBase
    {
        public override string Name => TableNames.Gsub;
        public override bool IsEssentialTable => false;

        /// <summary>
        /// Major version of the GSUB table. Set to 1 for current specification.
        /// </summary>
        public ushort MajorVersion { get; set; } = 1;

        /// <summary>
        /// Minor version of the GSUB table. 0 for GSUB 1.0, 1 for GSUB 1.1 (required for Variation Support).
        /// </summary>
        public ushort MinorVersion { get; set; } = 0; // Default to 1.0

        /// <summary>
        /// ScriptList table. Used for determining which language-specific features apply.
        /// </summary>
        public ScriptListTable ScriptList { get; set; }

        /// <summary>
        /// FeatureList table. Defines the specific typographic features (e.g., 'liga' for ligatures) 
        /// that are available in the font and links them to the Lookups that implement them.
        /// This list is referenced by ScriptList/LangSys to activate features for a given language.
        /// </summary>
        public FeatureListTable FeatureList { get; set; }

        public LookupListTable LookupList { get; set; }

        internal override void Clear()
        {
            throw new NotImplementedException();
        }

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

            // GSUB 1.1 FeatureVariations
            if (this.MajorVersion == 1 && this.MinorVersion == 1)
                writer.WriteUInt32BigEndian(0);

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

        /// <summary>
        /// Rewrites the GSUB table for font subsetting.
        /// </summary>
        /// <param name="context">The subsetting context containing glyph mappings.</param>
        /// <returns>A new GsubTable containing only the relevant substitutions.</returns>
        public GsubTable Rewrite(FontSubsettingContext context)
        {
            var newGsub = new GsubTable();
            newGsub.MajorVersion = this.MajorVersion;
            newGsub.MinorVersion = this.MinorVersion;

            // 1. Rewrite Lookups first - creates lookup index mapping
            LookupRewriteResult lookupResult = null;
            if (this.LookupList != null)
            {
                lookupResult = this.LookupList.Rewrite(context);
                if (lookupResult == null || lookupResult.NewLookupList == null || lookupResult.NewLookupList.Lookups.Count == 0)
                {
                    // No lookups remain - return null or minimal table
                    return null;
                }
                newGsub.LookupList = lookupResult.NewLookupList;
            }

            // 2. Rewrite FeatureList - pass lookup index mapping, get feature index mapping back
            FeatureRewriteResult featureResult = null;
            if (this.FeatureList != null && lookupResult != null)
            {
                featureResult = this.FeatureList.Rewrite(context, lookupResult.OldToNewIndexMap);
                if (featureResult == null || featureResult.NewFeatureList == null || featureResult.NewFeatureList.FeatureRecords.Count == 0)
                {
                    // No features remain
                    return null;
                }
                newGsub.FeatureList = featureResult.NewFeatureList;
            }

            // 3. Rewrite ScriptList - pass feature index mapping
            if (this.ScriptList != null && featureResult != null)
            {
                newGsub.ScriptList = this.ScriptList.Rewrite(context, featureResult.OldToNewIndexMap);
            }

            return newGsub;
        }
    }
}
