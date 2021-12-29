/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables
{
    public abstract class TableLoader<T>
        where T : class
    {
        public TableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables, string tableName)
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

        protected BigEndianBinaryReader _reader;
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
