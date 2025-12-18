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
 *************************************************************************************************/
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

        public void AddOrReplace(string key, object val)
        {
            if (_cachedTables.ContainsKey(key))
            {
                _cachedTables.Remove(key);
            }
            _cachedTables[key] = val;
        }

        public void Clear()
        {
            _cachedTables.Clear();
        }

        public int Count => _cachedTables.Keys.Count;
    }
}
