using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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
                UpdateOffsetAndSerialize(writer, tableStartOffset, scriptListOffPos, this.ScriptList);
            }

            // 2. FeatureList
            if (this.FeatureList != null)
            {
                UpdateOffsetAndSerialize(writer, tableStartOffset, featureListOffPos, this.FeatureList);
            }

            // 3. LookupList
            if (this.LookupList != null)
            {
                UpdateOffsetAndSerialize(writer, tableStartOffset, lookupListOffPos, this.LookupList);
            }
        }

        // Hjälpmetod för att hålla koden ren
        private void UpdateOffsetAndSerialize(FontsBinaryWriter writer, long tableStart, long placeholderPos, FontTableElement element)
        {
            ushort offset = (ushort)(writer.BaseStream.Position - tableStart);
            long resumePos = writer.BaseStream.Position;

            writer.BaseStream.Seek(placeholderPos, SeekOrigin.Begin);
            writer.WriteUInt16BigEndian(offset);

            writer.BaseStream.Seek(resumePos, SeekOrigin.Begin);
            element.Serialize(writer);
        }
    }
}
