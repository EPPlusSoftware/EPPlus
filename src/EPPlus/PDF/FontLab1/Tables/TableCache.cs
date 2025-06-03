using System.Collections.Generic;

namespace FontLab1.Tables
{
    internal static class TableCache
    {
        private static Dictionary<string, object> _cachedTables = new Dictionary<string, object>();

        public static object Get(string key)
        {
            return _cachedTables[key];
        }

        public static bool Contains(string key)
        {
            return _cachedTables.ContainsKey(key);
        }

        public static void Add(string key, object val)
        {
            _cachedTables.Add(key, val);
        }

        public static void Clear()
        {
            _cachedTables.Clear();
        }

        public static int Count => _cachedTables.Keys.Count;
    }
}
