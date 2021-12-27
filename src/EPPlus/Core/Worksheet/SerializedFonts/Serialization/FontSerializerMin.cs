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
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts.Serialization
{
    internal static class FontSerializerMin
    {
        public static readonly Encoding FileEncoding = Encoding.UTF8;
       

        public static SerializedFontMetrics Deserialize(Stream stream)
        {
            using (var reader = new BinaryReader(stream, FileEncoding))
            {
                var font = new SerializedFontMetrics
                {
                    Family = (SerializedFontFamilies)reader.ReadUInt16(),
                    SubFamily = (FontSubFamilies)reader.ReadUInt16(),
                    LineHeight = reader.ReadInt16(),
                    UnitsPerEm = reader.ReadUInt16(),
                    DefaultAdvanceWidth = reader.ReadInt16(),
                    AdvanceWidths = new Dictionary<char, short>()
                };
                var nRecords = reader.ReadInt16();
                for (var x = 0; x < nRecords; x++)
                {
                    var cc = reader.ReadUInt16();
                    var c = Convert.ToChar(cc);
                    var w = reader.ReadInt16();
                    font.AdvanceWidths[c] = w;
                }
                font.NumberOfKerningPairs = reader.ReadUInt16();
                if(font.NumberOfKerningPairs > 0)
                {
                    var pairs = new Dictionary<string, short>();
                    for(var x = 0; x < font.NumberOfKerningPairs; x++)
                    {
                        var l = reader.ReadUInt16();
                        var r = reader.ReadUInt16();
                        var cl = Convert.ToChar(l);
                        var cr = Convert.ToChar(r);
                        var pair = new KerningPair(cl, cr);
                        pairs.Add($"{pair.left}.{pair.right}", reader.ReadInt16());
                    }
                    font.KerningPairs = pairs;
                }
                return font;
            }
        }
    }
}
