using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class ExporterContext
    {
        //internal readonly Dictionary<string, int> _styleCache = new Dictionary<string, int>();
        //internal readonly Dictionary<string, int> _dxfStyleCache = new Dictionary<string, int>();

        internal readonly StyleCache _styleCache = new StyleCache();
        internal readonly StyleCache _dxfStyleCache = new StyleCache();

        internal ExporterContext() 
        {
            //_styleCache = new Dictionary<string, int>();
            //_dxfStyleCache = new Dictionary<string, int>();
        }


        //If multiple caches later perhaps enum cache type or simply a list with ids prefered over boolean.
        internal bool AddPairToCache(string key, int value, bool isDxfCache = false) 
        {
            if(isDxfCache) 
            {
                if(!_dxfStyleCache.ContainsKey(key))
                {
                    _dxfStyleCache.Add(key, value);
                    return true;
                }
            }
            else
            {
                if(!_styleCache.ContainsKey(key))
                {
                    _styleCache.Add(key, value);
                    return true;
                }
            }

            return false;
        }

        //If multiple caches later perhaps enum cache type or simply a list with ids prefered over boolean.
        internal int GetCacheId(string key, bool isDxfCache = false)
        {
            if (isDxfCache)
            {
                if (_dxfStyleCache.ContainsKey(key))
                {
                    return _dxfStyleCache[key];
                }
            }
            else
            {
                if (_styleCache.ContainsKey(key))
                {
                    return _styleCache[key];
                }
            }

            return -1;
        }

    }
}
