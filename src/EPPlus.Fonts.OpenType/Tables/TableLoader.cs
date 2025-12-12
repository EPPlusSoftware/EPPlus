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
using System;
using System.Threading;

namespace EPPlus.Fonts.OpenType.Tables
{
    internal abstract class TableLoader<T>
        where T : FontTableBase
    {
        //internal bool TableExists { get; private set; } = true;

        public TableLoader(TableLoaderSettings tblSettings, string tableName)
        {
            _reader = tblSettings._readerRef;
            if(tblSettings._tableRecordsRef.ContainsKey(tableName))
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

        public static object _syncRoot = new object();
        private bool _initialized;


        private bool _isLoading;
        private bool _isLoaded;

        public T Load(bool useCache = true)
        {
            lock (_syncRoot)
            {
                // If already loaded and cache is enabled
                if (_isLoaded && tableCache != null && tableCache.Contains(_tableName) && useCache)
                {
                    return tableCache.Get(_tableName) as T;
                }

                // If another thread is loading, wait until it's done
                while (_isLoading && !_isLoaded)
                {
                    Monitor.Wait(_syncRoot);
                }

                // If loaded after waiting, return from cache
                if (_isLoaded && tableCache != null)
                {
                    return tableCache.Get(_tableName) as T;
                }

                // Mark as loading
                _isLoading = true;

                // Set stream position under lock
                _reader.BaseStream.Position = _offset;

                // Load the table
                var t = LoadInternal();

                // Add to cache
                if (tableCache != null && !tableCache.Contains(_tableName))
                {
                    tableCache.Add(_tableName, t);
                }

                // Mark as loaded and notify waiting threads
                _isLoaded = true;
                _isLoading = false;
                Monitor.PulseAll(_syncRoot);

                return t;
            }
        }




        public void SetTable(string tableName, T value)
        {
            tableCache.AddOrReplace(tableName, value);
        }
    }
}
