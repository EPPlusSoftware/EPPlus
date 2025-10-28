using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables
{
    internal class TableCache
    {
        private Dictionary<string, object> _cachedTables = new Dictionary<string, object>();

        public object Get(string key)
        {
            return _cachedTables[key];
        }

        public bool Contains(string key)
        {
            return _cachedTables.ContainsKey(key);
        }

        public void Add(string key, object val)
        {
            _cachedTables.Add(key, val);
        }

        public void Clear()
        {
            _cachedTables.Clear();
        }

        public int Count => _cachedTables.Keys.Count;
    }
}
