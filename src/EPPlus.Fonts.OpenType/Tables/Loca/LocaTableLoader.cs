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
using EPPlus.Fonts.OpenType.Tables.Head;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Loca
{
    internal class LocaTableLoader : TableLoader<LocaTable>
    {
        private TableLoaderSettings LoadingSettingsRef;

        public LocaTableLoader(TableLoaderSettings tblSettings) : base(tblSettings, TableNames.Loca)
        {
            LoadingSettingsRef = tblSettings;
        }

        protected override LocaTable LoadInternal()
        {
            var headTable = TableLoaders.GetHeadTableLoader(LoadingSettingsRef).Load();
            var maxpTable = TableLoaders.GetMaxpTableLoader(LoadingSettingsRef).Load();
            _reader.BaseStream.Position = _offset;
            var indexes = new List<uint>();
            if(headTable.IndexToLocFormat == HeadTable.IndexToLocFormats.Offset16)
            {
                for(var x = 0; x < maxpTable.numGlyphs + 1; x++)
                {
                    var ix = Convert.ToUInt32(_reader.ReadUInt16BigEndian());
                    ix *= 2;
                    indexes.Add(ix);
                }
            }
            else if(headTable.IndexToLocFormat == HeadTable.IndexToLocFormats.Offset32)
            {
                for(var x = 0; x < maxpTable.numGlyphs + 1; x++)
                {
                    var ix = _reader.ReadUInt32BigEndian();
                    indexes.Add(ix);
                } 
            }
            return new LocaTable(maxpTable)
            {
                Offsets = indexes,
                IndexToLocFormat = headTable.IndexToLocFormat
            };
        }
    }
}
