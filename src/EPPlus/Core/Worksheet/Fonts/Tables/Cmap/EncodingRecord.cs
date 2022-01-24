/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Cmap
{
    public class EncodingRecord
    {
        public EncodingRecord(BigEndianBinaryReader reader)
        {
            _reader = reader;
            PlatformId = (Platforms)reader.ReadUInt16BigEndian();
            EncodingId = reader.ReadUInt16BigEndian();
            SubtableOffset = reader.ReadUInt32BigEndian();
        }

        private readonly BigEndianBinaryReader _reader;
        
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

        public GlyphMapping[] Mappings { get; set; }

        public IDictionary<ushort, char> GlyphIndexToCharMappings { get; internal set; }
    }
}
