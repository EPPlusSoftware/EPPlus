/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using System;
using System.IO;
using System.Text;
using EPPlus.Graphics;

namespace EPPlus.Export.Pdf.PdfFontData
{
    internal class PdfFontDataReader
    {
        public static readonly Encoding FileEncoding = Encoding.UTF8;

        public static PdfFontProperties Deserialize(Stream stream)
        {
            using (var reader = new BinaryReader(stream, FileEncoding))
            {
                var metrics = new PdfFontProperties();
                metrics.Version = reader.ReadUInt16();
                metrics.Family = (FontMetricsFamilies)reader.ReadUInt16();
                metrics.SubFamily = (FontSubFamilies)reader.ReadUInt16();
                metrics.LineHeight1em = reader.ReadSingle();
                metrics.DefaultWidthClass = (FontMetricsClass)reader.ReadByte();
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
                    var cls = (FontMetricsClass)reader.ReadByte();
                    var nRanges = reader.ReadUInt16();
                    for (var rngIx = 0; rngIx < nRanges; rngIx++)
                    {
                        var start = reader.ReadUInt16();
                        var end = reader.ReadUInt16();
                        for (var c = start; c <= end; c++)
                        {
                            metrics.CharMetrics[Convert.ToChar(c)] = cls;
                        }
                    }
                    var nCharactersInClass = reader.ReadUInt16();
                    if (nCharactersInClass == 0) continue;
                    for (int y = 0; y < nCharactersInClass; y++)
                    {
                        var cCode = reader.ReadUInt16();
                        var c = Convert.ToChar(cCode);
                        metrics.CharMetrics[c] = cls;
                    }
                }
                //we use height and width as xmax and ymax instead of actual width and height.
                metrics.FontBBox = new Rect();
                metrics.FontBBox.Width = reader.ReadInt16();
                metrics.FontBBox.X = reader.ReadInt16();
                metrics.FontBBox.Height = reader.ReadInt16();
                metrics.FontBBox.Y = reader.ReadInt16();
                metrics.ItalicAngle = reader.ReadDouble();
                metrics.Ascent = reader.ReadInt16();
                metrics.Descent = reader.ReadInt16();
                metrics.Capheight = reader.ReadInt16();
                metrics.Flags = reader.ReadInt32();
                metrics.StemV = reader.ReadDouble();
                metrics.FirstChar = 32;
                metrics.LastChar = 255;
                return metrics;
            }
        }
    }
}
