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
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.FontLocalization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Name
{
    public class NameTableLoader : TableLoader<NameTable>
    {
        public NameTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Name)
        {
        }

        protected override NameTable LoadInternal()
        {
            ushort format = _reader.ReadUInt16BigEndian();
            ushort count = _reader.ReadUInt16BigEndian();
            ushort stringOffset = _reader.ReadUInt16BigEndian();
            var globalStringOffset = _offset + stringOffset; 
            var records = new List<NameRecord>();
            for(var x = 0; x < count; x++)
            {
                var platformId = _reader.ReadUInt16BigEndian();
                var encodingID = _reader.ReadUInt16BigEndian();
                var languageID = _reader.ReadUInt16BigEndian();
                var nameID = _reader.ReadUInt16BigEndian();
                var length = _reader.ReadUInt16BigEndian();
                var offset = _reader.ReadUInt16BigEndian();
                var record = new NameRecord
                {
                    platformId = platformId,
                    encodingId = encodingID,
                    languageID = languageID,
                    nameId = nameID,
                    RecordType = (NameRecordTypes)nameID,
                    length = length,
                    offset = offset
                };
                records.Add(record);
            }
           
            foreach (var record in records)
            {
                _reader.BaseStream.Position = globalStringOffset + record.offset;
                var bytes = _reader.ReadBytes(record.length);
                
                // Macintosh platform
                if(record.platformId == 1)
                {
                    var enc = EncodingProviderCompatUtil.GetEncoding(10000);
                    record.Name = enc.GetString(bytes);
                    record.LanguageMapping = MacintoshLanguageMappings.Mappings[record.languageID];
                }
                else if(record.platformId == 0)
                {
                    record.Name = EncodingProviderCompatUtil.GetEncoding("utf-16BE").GetString(bytes);
                }
                // Windows platform
                else if(record.platformId == 3)
                {
                    if(WindowsLanguageMappings.Mappings.ContainsKey((int)record.languageID))
                    {
                        record.LanguageMapping = WindowsLanguageMappings.Mappings[record.languageID];
                    }
                    record.Name = EncodingProviderCompatUtil.GetEncoding("utf-16BE").GetString(bytes);
                }
            }
            
            return new NameTable
            {
                format = format,
                count = count,
                stringOffset = stringOffset,
                NameRecords = records.ToArray()
            };
        }

        private Encoding GetWindowsEncoding(int encoding)
        {
#if NETFULL
            switch(encoding)
            {
                case 0:
                    return Encoding.GetEncoding(1038);
                case 1:
                    return Encoding.GetEncoding("utf-16BE");
                 case 2:
                    return Encoding.GetEncoding("Shift-JIS");
                case 4:
                    return Encoding.GetEncoding("Big5");
                case 6:
                    return Encoding.GetEncoding("Johab");
                default:
                    return Encoding.Unicode;
            }
#else
            var encPrv = CodePagesEncodingProvider.Instance;
            switch(encoding)
            {
                case 0:
                    return encPrv.GetEncoding(1038);
                case 1:
                    return Encoding.GetEncoding("utf-16BE");
                case 2:
                    return encPrv.GetEncoding("Shift-JIS");
                case 4:
                    return encPrv.GetEncoding("Big5");
                case 6:
                    return encPrv.GetEncoding("Johab");
                default:
                    return Encoding.Unicode;
            }
#endif
        }
    }
}
