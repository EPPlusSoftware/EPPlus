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
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Cmap;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Glyph;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Head;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Hhea;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Hmtx;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Kern;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Loca;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Maxp;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Name;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Os2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables
{
    internal static class TableLoaders
    {
        public static LocaTableLoader GetLocaTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new LocaTableLoader(reader, tables);
        }

        public static HeadTableLoader GetHeadTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new HeadTableLoader(reader, tables);
        }

        public static CmapTableLoader GetCmapTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new CmapTableLoader(reader, tables);
        }

        public static GlyphTableLoader GetGlyphTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new GlyphTableLoader(reader, tables);
        }

        public static Os2TableLoader GetOs2TableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new Os2TableLoader(reader, tables);
        }

        public static HheaTableLoader GetHheaTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new HheaTableLoader(reader, tables);
        }

        public static MaxpTableLoader GetMaxpTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new MaxpTableLoader(reader, tables);
        }

        public static HmtxTableLoader GetHtmxTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new HmtxTableLoader(reader, tables);
        }

        public static NameTableLoader GetNameTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new NameTableLoader(reader, tables);
        }

        public static KernTableLoader GetKernTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables)
        {
            return new KernTableLoader(reader, tables);
        }
    }
}
