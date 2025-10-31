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
namespace EPPlus.Fonts.OpenType.Tables.Post
{
    internal class PostTableLoader : TableLoader<PostTable>
    {
        public PostTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Post)
        {
        }

        protected override PostTable LoadInternal()
        {
            var version = _reader.ReadInt32BigEndian();
            var italicMajor = _reader.ReadInt16BigEndian();
            var italicMinor = _reader.ReadUInt16BigEndian();
            var italicAngle = italicMajor + (italicMinor / 65536.0);
            var underlinePosition = _reader.ReadInt16BigEndian();
            var underlineThickness = _reader.ReadInt16BigEndian();
            var isFixedPitch = _reader.ReadUInt32BigEndian();
            var minMemType42 = _reader.ReadUInt32BigEndian();
            var maxMemType42 = _reader.ReadUInt32BigEndian();
            var minMemType1 = _reader.ReadUInt32BigEndian();
            var maxMemType1 = _reader.ReadUInt32BigEndian();

            return new PostTable()
            {
                version = version,
                italicAngle = italicAngle,
                underlinePosition = underlinePosition,
                underlineThickness = underlineThickness,
                isFixedPitch = isFixedPitch,
                minMemType42 = minMemType42,
                maxMemType42 = maxMemType42,
                minMemType1 = minMemType1,
                maxMemType1 = maxMemType1,
            };
        }
    }
}
