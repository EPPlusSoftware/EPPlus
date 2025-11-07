using System.Collections.Generic;

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
namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class EncodingRecord : FontTableElement
    {

        internal EncodingRecord(Platforms platformId, ushort encodingId, uint subtableOffset)
        {
            PlatformId = platformId;
            EncodingId = encodingId;
            SubtableOffset = subtableOffset;
        }


        internal EncodingRecord(FontsBinaryReader reader)
        {
            PlatformId = (Platforms)reader.ReadUInt16BigEndian();
            EncodingId = reader.ReadUInt16BigEndian();
            SubtableOffset = reader.ReadUInt32BigEndian();
        }
        
        /// <summary>
        /// 0 - Unicode
        /// 1 - Macintosh
        /// 2 - ISO (deprecated)
        /// 3 - Windows
        /// 4 - Custom
        /// </summary>
        public Platforms PlatformId { get; private set; }

       
        public ushort EncodingId { get; private set; }

        public uint SubtableOffset { get; set; }

        public CmapSubtableBase Subtable { get; internal set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt16BigEndian((ushort)PlatformId);
            writer.WriteUInt16BigEndian(EncodingId);
            writer.WriteUInt32BigEndian(SubtableOffset);
        }
    }
}
