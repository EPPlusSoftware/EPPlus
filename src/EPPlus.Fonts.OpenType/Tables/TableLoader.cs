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
  01/07/2026         EPPlus Software AB           Fixed threading with shared loaders
 *************************************************************************************************/
using System.Collections.Generic;
using System;
using System.Threading;

namespace EPPlus.Fonts.OpenType.Tables
{
    internal abstract class TableLoader<T>
        where T : FontTableBase
    {
        public TableLoader(TableLoaderSettings tblSettings, string tableName)
        {
            _reader = tblSettings._readerRef;
            if (tblSettings._tableRecordsRef.ContainsKey(tableName))
            {
                _offset = tblSettings._tableRecordsRef[tableName].Offset;
                _length = tblSettings._tableRecordsRef[tableName].Length;
            }
            _tables = tblSettings._tableRecordsRef;
            _tableName = tableName;
            tableCache = tblSettings._tblCacheRef;
        }

        protected FontsBinaryReader _reader;
        private readonly string _tableName;
        protected readonly uint _offset;
        protected readonly uint _length;
        protected Dictionary<string, TableRecord> _tables;
        internal TableCache tableCache;

        protected abstract T LoadInternal();

        // ✅ Instance lock for this specific table loader
        private readonly object _instanceLock = new object();

        private bool _isLoading;
        private bool _isLoaded;

        public T Load(bool useCache = true)
        {
            // First: Check cache with instance lock (fast path)
            lock (_instanceLock)
            {
                if (_isLoaded && tableCache != null && tableCache.Contains(_tableName) && useCache)
                {
                    return tableCache.Get(_tableName) as T;
                }

                while (_isLoading && !_isLoaded)
                {
                    Monitor.Wait(_instanceLock);
                }

                if (_isLoaded && tableCache != null)
                {
                    return tableCache.Get(_tableName) as T;
                }

                _isLoading = true;
            }

            T t;

            // Second: Load table with reader lock
            lock (_reader)
            {
                // ✅ CRITICAL: Save and restore stream position!
                long savedPosition = _reader.BaseStream.Position;

                try
                {
                    // Seek to table start
                    _reader.BaseStream.Position = _offset;

                    // Load the table
                    t = LoadInternal();
                }
                finally
                {
                    // ✅ RESTORE position after loading
                    // This prevents interfering with other loaders
                    _reader.BaseStream.Position = savedPosition;
                }
            }

            // Third: Update cache
            lock (_instanceLock)
            {
                if (tableCache != null && !tableCache.Contains(_tableName))
                {
                    tableCache.Add(_tableName, t);
                }

                _isLoaded = true;
                _isLoading = false;
                Monitor.PulseAll(_instanceLock);
            }

            return t;
        }

        public void SetTable(string tableName, T value)
        {
            tableCache.AddOrReplace(tableName, value);
        }
    }
}