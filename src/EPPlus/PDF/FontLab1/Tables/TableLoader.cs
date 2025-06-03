using System.Collections.Generic;

namespace FontLab1.Tables
{
    internal abstract class TableLoader<T>
        where T : class
    {
        public TableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables, string tableName)
        {
            _reader = reader;
            if(tables.ContainsKey(tableName))
            {
                _offset = tables[tableName].Offset;
                _length = tables[tableName].Length;
            }
            _tables = tables;
            _tableName = tableName;
            _reader.BaseStream.Position = _offset;
        }

        protected MyBinaryReader _reader;
        private readonly string _tableName;
        protected readonly uint _offset;
        protected readonly uint _length;
        protected Dictionary<string, TableRecord> _tables;
        private static Dictionary<string, object> _cachedTables = new Dictionary<string, object>();

        protected abstract T LoadInternal();

        public T Load(bool useCache = true)
        {
            if(TableCache.Contains(_tableName) && useCache)
            {
                return TableCache.Get(_tableName) as T;
            }
            else if(!TableCache.Contains(_tableName))
            {
                _reader.BaseStream.Position = _offset;
                var t = LoadInternal();
                TableCache.Add(_tableName, t);
                return t;
            }
            else
            {
                return default(T);
            }
        }
    }
}
