using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Glyph;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Hhea;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Hmtx;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Kern;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Loca;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Maxp;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Name;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Os2;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Post;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Head;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Cmap;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables
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
