using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Kern;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables
{
    internal static class TableLoaders
    {
        public static LocaTableLoader GetLocaTableLoader(TableLoaderSettings settings)
        {
            return new LocaTableLoader(settings);
        }

        public static HeadTableLoader GetHeadTableLoader(TableLoaderSettings settings)
        {
            return new HeadTableLoader(settings);
        }

        public static CmapTableLoader GetCmapTableLoader(TableLoaderSettings settings)
        {
            return new CmapTableLoader(settings);
        }

        public static GlyphTableLoader GetGlyphTableLoader(TableLoaderSettings settings)
        {
            return new GlyphTableLoader(settings);
        }

        public static Os2TableLoader GetOs2TableLoader(TableLoaderSettings settings)
        {
            return new Os2TableLoader(settings);
        }

        public static HheaTableLoader GetHheaTableLoader(TableLoaderSettings settings)
        {
            return new HheaTableLoader(settings);
        }

        public static MaxpTableLoader GetMaxpTableLoader(TableLoaderSettings settings)
        {
            return new MaxpTableLoader(settings);
        }

        public static HmtxTableLoader GetHtmxTableLoader(TableLoaderSettings settings)
        {
            return new HmtxTableLoader(settings);
        }

        public static NameTableLoader GetNameTableLoader(TableLoaderSettings settings)
        {
            return new NameTableLoader(settings);
        }

        public static KernTableLoader GetKernTableLoader(TableLoaderSettings settings)
        {
            return new KernTableLoader(settings);
        }

        public static PostTableLoader GetPostTableLoader(TableLoaderSettings settings)
        {
            return new PostTableLoader(settings);
        }
    }
}
