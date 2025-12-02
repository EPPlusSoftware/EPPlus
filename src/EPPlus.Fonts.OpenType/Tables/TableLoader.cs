using System.Collections.Generic;
using System;

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
            _reader.BaseStream.Position = _offset;
            tableCache = tblSettings._tblCacheRef;
        }

        protected FontsBinaryReader _reader;
        private readonly string _tableName;
        protected readonly uint _offset;
        protected readonly uint _length;
        protected Dictionary<string, TableRecord> _tables;
        private static Dictionary<string, object> _cachedTables = new Dictionary<string, object>();
        internal TableCache tableCache;

        protected abstract T LoadInternal();

        public static object _syncRoot = new object();
        public T Load(bool useCache = true)
        {
            lock (_syncRoot)
            {
                if (tableCache != null && tableCache.Contains(_tableName) && useCache)
                {
                    return tableCache.Get(_tableName) as T;
                }
                else if (tableCache == null || !tableCache.Contains(_tableName))
                {
                    _reader.BaseStream.Position = _offset;
                    var t = LoadInternal();

                    if (tableCache != null)
                    {
                        tableCache.Add(_tableName, t);
                    }

                    return t;
                }
                else
                {
                    return default(T);
                }
            }
        }

        public void SetTable(string tableName, T value)
        {
            tableCache.AddOrReplace(tableName, value);
        }
    }
}
