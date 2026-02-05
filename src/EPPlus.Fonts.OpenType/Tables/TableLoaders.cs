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
  01/14/2026         EPPlus Software AB           Cache loader instances for thread safety
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
    internal static class TableLoaders
    {
        public static LocaTableLoader GetLocaTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.LocaLoader == null)
                {
                    cache.LocaLoader = new LocaTableLoader(settings);
                }
                return cache.LocaLoader;
            }
        }

        public static HeadTableLoader GetHeadTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.HeadLoader == null)
                {
                    cache.HeadLoader = new HeadTableLoader(settings);
                }
                return cache.HeadLoader;
            }
        }

        public static CmapTableLoader GetCmapTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.CmapLoader == null)
                {
                    cache.CmapLoader = new CmapTableLoader(settings);
                }
                return cache.CmapLoader;
            }
        }

        public static GlyfTableLoader GetGlyfTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.GlyfLoader == null)
                {
                    cache.GlyfLoader = new GlyfTableLoader(settings);
                }
                return cache.GlyfLoader;
            }
        }

        public static Os2TableLoader GetOs2TableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.Os2Loader == null)
                {
                    cache.Os2Loader = new Os2TableLoader(settings);
                }
                return cache.Os2Loader;
            }
        }

        public static HheaTableLoader GetHheaTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.HheaLoader == null)
                {
                    cache.HheaLoader = new HheaTableLoader(settings);
                }
                return cache.HheaLoader;
            }
        }

        public static MaxpTableLoader GetMaxpTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.MaxpLoader == null)
                {
                    cache.MaxpLoader = new MaxpTableLoader(settings);
                }
                return cache.MaxpLoader;
            }
        }

        public static HmtxTableLoader GetHmtxTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.HmtxLoader == null)
                {
                    cache.HmtxLoader = new HmtxTableLoader(settings);
                }
                return cache.HmtxLoader;
            }
        }

        public static NameTableLoader GetNameTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.NameLoader == null)
                {
                    cache.NameLoader = new NameTableLoader(settings);
                }
                return cache.NameLoader;
            }
        }

        public static KernTableLoader GetKernTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.KernLoader == null)
                {
                    cache.KernLoader = new KernTableLoader(settings);
                }
                return cache.KernLoader;
            }
        }

        public static PostTableLoader GetPostTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.PostLoader == null)
                {
                    cache.PostLoader = new PostTableLoader(settings);
                }
                return cache.PostLoader;
            }
        }

        public static GsubTableLoader GetGsubTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.GsubLoader == null)
                {
                    cache.GsubLoader = new GsubTableLoader(settings);
                }
                return cache.GsubLoader;
            }
        }

        public static GposTableLoader GetGposTableLoader(TableLoaderSettings settings)
        {
            var cache = settings._loaderCacheRef;
            lock (cache.SyncLock)
            {
                if (cache.GposLoader == null)
                {
                    cache.GposLoader = new GposTableLoader(settings);
                }
                return cache.GposLoader;
            }
        }
    }
}