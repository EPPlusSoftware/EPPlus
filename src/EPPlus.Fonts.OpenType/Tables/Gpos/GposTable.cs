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
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Features;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
            // Store table start position
            long tableStart = writer.BaseStream.Position;

            // Write version
            writer.WriteUInt16BigEndian(MajorVersion);
            writer.WriteUInt16BigEndian(MinorVersion);

            // Write placeholder offsets (will update later)
            long scriptListOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // ScriptList offset placeholder

            long featureListOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // FeatureList offset placeholder

            long lookupListOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // LookupList offset placeholder

            long featureVariationsOffsetPos = 0;
            if (MinorVersion >= 1)
            {
                featureVariationsOffsetPos = writer.BaseStream.Position;
                writer.WriteUInt32BigEndian(0); // FeatureVariations offset placeholder (not implemented)
            }

            // Write ScriptList
            if (ScriptList != null)
            {
                long currentPos = writer.BaseStream.Position;
                ushort scriptListOffset = (ushort)(currentPos - tableStart);

                // Go back and write actual offset
                writer.BaseStream.Seek(scriptListOffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(scriptListOffset);

                // Return to current position and serialize
                writer.BaseStream.Seek(currentPos, System.IO.SeekOrigin.Begin);
                ScriptList.Serialize(writer);
            }

            // Write FeatureList
            if (FeatureList != null)
            {
                long currentPos = writer.BaseStream.Position;
                ushort featureListOffset = (ushort)(currentPos - tableStart);

                writer.BaseStream.Seek(featureListOffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(featureListOffset);

                writer.BaseStream.Seek(currentPos, System.IO.SeekOrigin.Begin);
                FeatureList.Serialize(writer);
            }

            // Write LookupList
            if (LookupList != null)
            {
                long currentPos = writer.BaseStream.Position;
                ushort lookupListOffset = (ushort)(currentPos - tableStart);

                writer.BaseStream.Seek(lookupListOffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(lookupListOffset);

                writer.BaseStream.Seek(currentPos, System.IO.SeekOrigin.Begin);
                LookupList.Serialize(writer);
            }

            // Note: FeatureVariations not implemented yet
            // When implemented, write at featureVariationsOffsetPos
        }

        internal override void Clear()
        {
            // TODO: Implement when needed for memory management
            // This would clear internal caches/state if any
            ScriptList = null;
            FeatureList = null;
            LookupList = null;
        }
    }
}
