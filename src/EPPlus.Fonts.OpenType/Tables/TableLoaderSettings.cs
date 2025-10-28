using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables
{
    internal class TableLoaderSettings
    {
        internal FontsBinaryReader _readerRef { get; private set; }
        internal Dictionary<string, TableRecord> _tableRecordsRef { get; private set; }
        internal TableCache _tblCacheRef { get; private set; }

        internal TableLoaderSettings(FontsBinaryReader reader, Dictionary<string, TableRecord> records, TableCache tblCache) 
        {
            _readerRef = reader;
            _tableRecordsRef = records;
            _tblCacheRef = tblCache;
        }
    }
}
