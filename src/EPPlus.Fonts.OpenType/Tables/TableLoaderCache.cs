/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/14/2026         EPPlus Software AB           Loader cache for thread safety
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Kern;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;

namespace EPPlus.Fonts.OpenType.Tables
{
    /// <summary>
    /// Cache for TableLoader instances to ensure thread-safe access.
    /// Each OpenTypeFont instance has its own loader cache.
    /// </summary>
    internal class TableLoaderCache
    {
        private readonly object _lock = new object();

        public LocaTableLoader LocaLoader;
        public HeadTableLoader HeadLoader;
        public CmapTableLoader CmapLoader;
        public GlyfTableLoader GlyfLoader;
        public Os2TableLoader Os2Loader;
        public HheaTableLoader HheaLoader;
        public MaxpTableLoader MaxpLoader;
        public HmtxTableLoader HmtxLoader;
        public NameTableLoader NameLoader;
        public KernTableLoader KernLoader;
        public PostTableLoader PostLoader;
        public GsubTableLoader GsubLoader;
        public GposTableLoader GposLoader;

        internal object SyncLock { get { return _lock; } }
    }
}