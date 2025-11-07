using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
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
    internal class CmapTableLoader : TableLoader<CmapTable>
    {
        public CmapTableLoader(TableLoaderSettings settings) : base(settings, "cmap")
        {
        }

        protected override CmapTable LoadInternal()
        {
            var table = new CmapTable
            {
                Version = _reader.ReadUInt16BigEndian(),
                NumTables = _reader.ReadUInt16BigEndian()
            };

            for (var x = 0; x < table.NumTables; x++)
            {
                var enc = new EncodingRecord(_reader);
                table.EncodingRecords.Add(enc);
            }

            for (var x = 0; x < table.NumTables; x++)
            {
                var enc = table.EncodingRecords[x];
                var currentPos = _offset + enc.SubtableOffset;
                _reader.BaseStream.Position = currentPos;

                var format = _reader.ReadUInt16BigEndian();
                _reader.BaseStream.Position = currentPos; // rewind to start of subtable

                switch (format)
                {
                    case 0:
                        var sub0 = new CmapSubtable0Deserializer(_reader).Deserialize(currentPos);
                        table.SubTables.Add(sub0);
                        enc.Subtable = sub0;
                        break;

                    case 4:
                        var sub4 = new CmapSubtable4Deserializer(_reader).Deserialize(currentPos);
                        table.SubTables.Add(sub4);
                        enc.Subtable = sub4;
                        break;

                    case 6:
                        var sub6 = new CmapSubtable6Deserializer(_reader).Deserialize(currentPos);
                        table.SubTables.Add(sub6);
                        enc.Subtable = sub6;
                        break;

                    default:
                        // Optional: handle unsupported formats
                        break;
                }
            }

            return table;
        }


        //protected override CmapTable LoadInternal()
        //{
        //    var table = new CmapTable
        //    {
        //        Version = _reader.ReadUInt16BigEndian(),
        //        NumTables = _reader.ReadUInt16BigEndian()
        //    };

        //    for(var x = 0; x < table.NumTables; x++)
        //    {
        //        var enc = new EncodingRecord(_reader);
        //        table.EncodingRecords.Add(enc);
        //    }

        //    for(var x = 0; x < table.NumTables; x++)
        //    {
        //        var enc = table.EncodingRecords[x];
        //        var currentPos = _offset + enc.SubtableOffset;
        //        _reader.BaseStream.Position = currentPos;
        //        var format = _reader.ReadUInt16BigEndian();
        //        if(format == 0)
        //        {
        //            var subtable = new CmapSubtable0(_reader);
        //            table.SubTables.Add(subtable);
        //            enc.Subtable = subtable;
        //            //enc.Mappings = subtable.GlyphMappingArray;
        //        }
        //        else if(format == 4)
        //        {
        //            var subtable = new CmapSubtable4(_reader);
        //            table.SubTables.Add(subtable);
        //            enc.Subtable = subtable;
        //            //enc.Mappings = subtable.GlyphMappingArray;
        //            //enc.GlyphIndexToCharMappings = subtable.GlyphIndexToCharMappings;
        //            //enc.CharMappingsToGlyphIndex = subtable.CharMappingsToGlyphIndex;
        //        }
        //        else if(format == 6)
        //        {
        //            var subtable = new CmapSubtable6(_reader);
        //            table.SubTables.Add(subtable);
        //            enc.Subtable = subtable;
        //            //enc.Mappings = subtable.GlyphMappingArray;
        //        }
        //    }
        //    return table;
        //}
    }
}
