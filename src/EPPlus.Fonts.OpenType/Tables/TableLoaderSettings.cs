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
  01/14/2026         EPPlus Software AB           Added loader cache reference
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables
{
    internal class TableLoaderSettings
    {
        internal Dictionary<string, TableRecord> _tableRecordsRef { get; private set; }
        internal TableCache _tblCacheRef { get; private set; }
        internal TableLoaderCache _loaderCacheRef { get; private set; }

        internal FontTableReaderFactory TableReaderFactory { get; private set; }

        internal TableLoaderSettings(
            FontTableReaderFactory tableReaderFactory,
            Dictionary<string, TableRecord> records,
            TableCache tblCache,
            TableLoaderCache loaderCache)
        {
            TableReaderFactory = tableReaderFactory;
            _tableRecordsRef = records;
            _tblCacheRef = tblCache;
            _loaderCacheRef = loaderCache;
        }
    }
}