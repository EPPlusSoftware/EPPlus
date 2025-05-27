using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.PDF.PdfPageSettings;

namespace OfficeOpenXml.PDF.PdfFontData
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
                //read pdf data:
                var xmax = reader.ReadInt16();
                var xmin = reader.ReadInt16();
                var ymax = reader.ReadInt16();
                var ymin = reader.ReadInt16();
                //we use height and width as xmax and ymax instead of actual width and height.
                metrics.fontBBox = new PdfRect();
                metrics.fontBBox.X = xmin;
                metrics.fontBBox.Y = ymin;
                metrics.fontBBox.Width = xmax;
                metrics.fontBBox.Height = ymax;

                metrics.italicAngle = reader.ReadDouble();
                metrics.ascent = reader.ReadInt16();
                metrics.descent = reader.ReadInt16();
                metrics.capheight = reader.ReadInt16();
                metrics.flags = reader.ReadInt32();

                metrics.stemV = 1;
                metrics.firstChar = 1;
                metrics.lastChar = 1;
                

                return metrics;
            }
        }
    }
}
