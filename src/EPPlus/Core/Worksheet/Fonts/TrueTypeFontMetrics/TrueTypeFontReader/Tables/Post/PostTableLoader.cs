using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Post
{
    internal class PostTableLoader : TableLoader<PostTable>
    {
        public PostTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Post)
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
