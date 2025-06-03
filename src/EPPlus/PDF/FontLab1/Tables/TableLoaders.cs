using FontLab1.Tables.Cmap;
using FontLab1.Tables.Glyph;
using FontLab1.Tables.Head;
using FontLab1.Tables.Hhea;
using FontLab1.Tables.Hmtx;
using FontLab1.Tables.Kern;
using FontLab1.Tables.Loca;
using FontLab1.Tables.Maxp;
using FontLab1.Tables.Name;
using FontLab1.Tables.Os2;
using FontLab1.Tables.Post;
using System.Collections.Generic;

namespace FontLab1.Tables
{
    internal static class TableLoaders
    {
        public static LocaTableLoader GetLocaTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new LocaTableLoader(reader, tables);
        }

        public static HeadTableLoader GetHeadTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new HeadTableLoader(reader, tables);
        }

        public static CmapTableLoader GetCmapTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new CmapTableLoader(reader, tables);
        }

        public static GlyphTableLoader GetGlyphTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new GlyphTableLoader(reader, tables);
        }

        public static Os2TableLoader GetOs2TableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new Os2TableLoader(reader, tables);
        }

        public static HheaTableLoader GetHheaTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new HheaTableLoader(reader, tables);
        }

        public static MaxpTableLoader GetMaxpTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new MaxpTableLoader(reader, tables);
        }

        public static HmtxTableLoader GetHtmxTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new HmtxTableLoader(reader, tables);
        }

        public static NameTableLoader GetNameTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new NameTableLoader(reader, tables);
        }

        public static KernTableLoader GetKernTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new KernTableLoader(reader, tables);
        }

        public static PostTableLoader GetPostTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new PostTableLoader(reader, tables);
        }
    }
}
