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
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables
{
    internal class TableLoaders
    {
        public LocaTableLoader GetLocaTableLoader(TableLoaderSettings settings)
        {
            return new LocaTableLoader(settings);
        }

        public HeadTableLoader GetHeadTableLoader(TableLoaderSettings settings)
        {
            return new HeadTableLoader(settings);
        }

        public CmapTableLoader GetCmapTableLoader(TableLoaderSettings settings)
        {
            return new CmapTableLoader(settings);
        }

        public GlyfTableLoader GetGlyfTableLoader(TableLoaderSettings settings)
        {
            return new GlyfTableLoader(settings);
        }

        public Os2TableLoader GetOs2TableLoader(TableLoaderSettings settings)
        {
            return new Os2TableLoader(settings);
        }

        public HheaTableLoader GetHheaTableLoader(TableLoaderSettings settings)
        {
            return new HheaTableLoader(settings);
        }

        public MaxpTableLoader GetMaxpTableLoader(TableLoaderSettings settings)
        {
            return new MaxpTableLoader(settings);
        }

        public HmtxTableLoader GetHmtxTableLoader(TableLoaderSettings settings)
        {
            return new HmtxTableLoader(settings);
        }

        public NameTableLoader GetNameTableLoader(TableLoaderSettings settings)
        {
            return new NameTableLoader(settings);
        }

        public KernTableLoader GetKernTableLoader(TableLoaderSettings settings)
        {
            return new KernTableLoader(settings);
        }

        public PostTableLoader GetPostTableLoader(TableLoaderSettings settings)
        {
            return new PostTableLoader(settings);
        }

        public bool IsSupportedTable(string name)
        {
            switch(name)
            {
                case TableNames.Loca:
                case TableNames.Head:
                case TableNames.Hhea:
                case TableNames.Hmtx:
                case TableNames.Maxp:
                case TableNames.Name:
                case TableNames.Kern:
                case TableNames.Os2:
                case TableNames.Post:
                case TableNames.Cmap:
                case TableNames.Glyf:
                    return true;
                default:
                    return false;
            }
        }

        public FontTableBase GetTable(string name, TableLoaderSettings settings)
        {
            if(string.IsNullOrEmpty(name)) throw new ArgumentNullException("name");
            if(!IsSupportedTable(name.ToLower()))
            {
                throw new NotSupportedException("Not supported table: " + name);
            }
            switch(name)
            {
                case TableNames.Loca:
                    return GetLocaTableLoader(settings).Load();
                case TableNames.Head:
                    return GetHeadTableLoader(settings).Load();
                case TableNames.Hhea:
                    return GetHheaTableLoader(settings).Load();
                case TableNames.Hmtx:
                    return GetHmtxTableLoader(settings).Load();
                case TableNames.Maxp:
                    return GetMaxpTableLoader(settings).Load();
                case TableNames.Name:
                    return GetNameTableLoader(settings).Load();
                case TableNames.Kern:
                    return GetKernTableLoader(settings).Load();
                case TableNames.Os2:
                    return GetOs2TableLoader(settings).Load();
                case TableNames.Post:
                    return GetPostTableLoader(settings).Load();
                case TableNames.Glyf:
                    return GetGlyfTableLoader(settings).Load();
                case TableNames.Cmap:
                    return GetCmapTableLoader(settings).Load();
                default:
                    return null;

            }
        }
    }
}
