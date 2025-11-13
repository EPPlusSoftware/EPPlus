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
            var versionRaw = _reader.ReadInt32BigEndian();
            var italicAngleRaw = _reader.ReadInt32BigEndian();
            var underlinePosition = _reader.ReadInt16BigEndian();
            var underlineThickness = _reader.ReadInt16BigEndian();
            var isFixedPitch = _reader.ReadUInt32BigEndian();
            var minMemType42 = _reader.ReadUInt32BigEndian();
            var maxMemType42 = _reader.ReadUInt32BigEndian();
            var minMemType1 = _reader.ReadUInt32BigEndian();
            var maxMemType1 = _reader.ReadUInt32BigEndian();

            var post = new PostTable()
            {
                version = new Version16Dot16(versionRaw),
                italicAngle = new Fixed16Dot16(italicAngleRaw),
                underlinePosition = underlinePosition,
                underlineThickness = underlineThickness,
                isFixedPitch = isFixedPitch,
                minMemType42 = minMemType42,
                maxMemType42 = maxMemType42,
                minMemType1 = minMemType1,
                maxMemType1 = maxMemType1,
            };
            if(post.version.Major == 2 && post.version.Minor == 0)
            {
                post.numGlyphs = _reader.ReadUInt16BigEndian();
                post.glyphNameIndex = new List<ushort>();
                for(var i = 0; i < post.numGlyphs; i++)
                {
                    post.glyphNameIndex.Add(_reader.ReadUInt16BigEndian());
                }

                // Read the Pascal-strings
                var stringList = new List<string>();
                while (_reader.BaseStream.Position < _offset + _length)
                {
                    byte len = _reader.ReadByte();
                    var strBytes = _reader.ReadBytes(len);
                    var str = System.Text.Encoding.ASCII.GetString(strBytes);
                    stringList.Add(str);
                }

                post.glyphNames = stringList;

            }
            return post;
        }
    }
}
