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
    using System.Threading;

    internal class TableCache
    {
        private readonly Dictionary<string, object> _cachedTables = new Dictionary<string, object>();
        private readonly ReaderWriterLockSlim _lock = new ReaderWriterLockSlim();

        public bool TryGet(string key, out object value)
        {
            _lock.EnterReadLock();
            try
            {
                return _cachedTables.TryGetValue(key, out value);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        public object Get(string key)
        {
            _lock.EnterReadLock();
            try
            {
                return _cachedTables[key];
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        public bool Contains(string key)
        {
            _lock.EnterReadLock();
            try
            {
                return _cachedTables.ContainsKey(key);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        public void Add(string key, object val)
        {
            _lock.EnterWriteLock();
            try
            {
                _cachedTables.Add(key, val);
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        public void AddOrReplace(string key, object val)
        {
            _lock.EnterWriteLock();
            try
            {
                _cachedTables[key] = val;  // Enklare än ContainsKey + Remove
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        public void Clear()
        {
            _lock.EnterWriteLock();
            try
            {
                _cachedTables.Clear();
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        public int Count
        {
            get
            {
                _lock.EnterReadLock();
                try
                {
                    return _cachedTables.Count;
                }
                finally
                {
                    _lock.ExitReadLock();
                }
            }
        }
    }
}
