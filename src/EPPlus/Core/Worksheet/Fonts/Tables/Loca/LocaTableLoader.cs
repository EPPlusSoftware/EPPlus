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
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Head;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Loca
{
    public class LocaTableLoader : TableLoader<LocaTable>
    {
        public LocaTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Loca)
        {

        }

        protected override LocaTable LoadInternal()
        {
            var headTable = TableLoaders.GetHeadTableLoader(_reader, _tables).Load();
            var maxpTable = TableLoaders.GetMaxpTableLoader(_reader, _tables).Load();
            _reader.BaseStream.Position = _offset;
            var indexes = new List<uint>();
            if(headTable.IndexToLocFormat == HeadTable.IndexToLocFormats.Offset16)
            {
                for(var x = 0; x <= maxpTable.numGlyphs + 1; x++)
                {
                    var ix = Convert.ToUInt32(_reader.ReadUInt16BigEndian());
                    ix *= 2;
                    indexes.Add(ix);
                }
            }
            else if(headTable.IndexToLocFormat == HeadTable.IndexToLocFormats.Offset32)
            {
                for(var x = 0; x <= maxpTable.numGlyphs + 1; x++)
                {
                    var ix = _reader.ReadUInt32BigEndian();
                    indexes.Add(ix);
                }
                
            }
            return new LocaTable
            {
                Offsets = indexes.ToArray()
            };
        }
    }
}
