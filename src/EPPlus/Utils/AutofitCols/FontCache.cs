using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
#if !NET35
using System.Collections.Concurrent;
#endif

namespace OfficeOpenXml.Utils.AutofitCols
{
    internal class FontCache
    {
        private readonly object _syncRoot = new object();
#if NET35
        private Dictionary<int, Font> _fonts = new Dictionary<int, Font>();
        
#else
        private ConcurrentDictionary<int, Font> _fonts = new ConcurrentDictionary<int, Font>();
#endif



#if NET35
        internal Font this[int ix]
        {
            get
            {
                return _fonts[ix];
            }
            set
            {
                lock(_syncRoot)
                {
                    _fonts[ix] = value;
                }
            }
        }

        internal void Add(int ix, Font font)
        {
            lock(_syncRoot)
            {
                _fonts.Add(ix, font);
            }
        }
#else
        internal Font this[int ix]
        {
            get
            {
                return _fonts[ix];
            }
            set
            {
                _fonts[ix] = value;
            }
        }

        internal void Add(int key, Font font)
        {
            _fonts[key] = font;
        }
#endif

        internal bool ContainsKey(int key)
        {
            return _fonts.ContainsKey(key);
        }
    }
}
