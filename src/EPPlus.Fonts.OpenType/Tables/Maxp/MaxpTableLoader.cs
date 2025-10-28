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
namespace EPPlus.Fonts.OpenType.Tables.Maxp
{
    internal class MaxpTableLoader : TableLoader<MaxpTable>
    {
        public MaxpTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Maxp)
        {
        }

        protected override MaxpTable LoadInternal()
        {
            var pos = _reader.BaseStream.Position;
            var version = _reader.ReadInt32BigEndian();
            var major = (version >> 16);
            var minor = (version & 16);
            var pos2 = _reader.BaseStream.Position;
            var nGlyphs = _reader.ReadUInt16BigEndian();
            return new MaxpTable
            {
                numGlyphs = nGlyphs
            };
        }
    }
}
