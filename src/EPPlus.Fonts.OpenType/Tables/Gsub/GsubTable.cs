using System;
using System.Collections.Generic;
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

            // 0: Fixed Version (FWORD = 4 bytes total)
            // USHORT MajorVersion (1)
            writer.WriteUInt16BigEndian(this.MajorVersion);
            // USHORT MinorVersion (0 or 1)
            writer.WriteUInt16BigEndian(this.MinorVersion);

            // 4: USHORT ScriptListOffset
            long scriptListOffsetPosition = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian((ushort)0); // Placeholder

            // 6: USHORT FeatureListOffset
            // ... (resten av koden för offsets)
        }
    }
}
