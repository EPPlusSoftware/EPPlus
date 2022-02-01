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
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal static class GenericFontMetricsSerializer
    {
        public static readonly Encoding FileEncoding = Encoding.UTF8;

        public static SerializedFontMetrics Deserialize(Stream stream)
        {
            using (var reader = new BinaryReader(stream, FileEncoding))
            {
                var metrics = new SerializedFontMetrics();
                metrics.Version = reader.ReadUInt16();
                metrics.Family = (FontMetricsFamilies)reader.ReadUInt16();
                metrics.SubFamily = (FontSubFamilies)reader.ReadUInt16();
                metrics.LineHeight1em = reader.ReadSingle();
                metrics.DefaultWidth1em = reader.ReadSingle();
                var nClassWidths = reader.ReadUInt16();
                if (nClassWidths == 0)
                {
                    return metrics;
                }
                for (var x = 0; x < nClassWidths; x++)
                {
                    var cls = (FontMetricsClass)reader.ReadByte();
                    var width = reader.ReadSingle();
                    metrics.ClassWidths[cls] = width;
                }
                var nClasses = reader.ReadUInt16();
                for (var x = 0; x < nClasses; x++)
                {
                    var nCharactersInClass = reader.ReadUInt16();
                    if (nCharactersInClass == 0) continue;
                    var cls = (FontMetricsClass)reader.ReadByte();
                    for (int y = 0; y < nCharactersInClass; y++)
                    {
                        var cCode = reader.ReadUInt16();
                        var c = Convert.ToChar(cCode);
                        metrics.CharMetrics[c] = cls;
                    }
                }
                return metrics;
            }
        }
    }
}
