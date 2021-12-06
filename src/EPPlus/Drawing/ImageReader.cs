/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  12/03/2021         EPPlus Software AB       Added
 *************************************************************************************************/
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using OfficeOpenXml.Packaging.Ionic.Zip;
namespace OfficeOpenXml.Drawing
{
    internal class ImageReader
    {
        internal struct TifIfd
        {
            public short Tag;
            public short Type;
            public int Count;
            public int ValueOffset;
        }
        internal static bool TryGetImageBounds(ePictureType pictureType, MemoryStream ms, ref double width, ref double height)
        {
            width = 0;
            height = 0;
            try
            {
                ms.Seek(0, SeekOrigin.Begin);
                if (pictureType == ePictureType.Png && IsPng(ms, ref width, ref height))
                {
                    return true;
                }
                if (pictureType == ePictureType.Emf && IsEmf(ms, ref width, ref height))
                {
                    return true;
                }
                if (pictureType == ePictureType.Wmf && IsWmf(ms, ref width, ref height))
                {
                    return true;
                }
                else if (pictureType == ePictureType.Svg && IsSvg(ms, ref width, ref height))
                {
                    return true;
                }
                else if (pictureType == ePictureType.Tif && IsTif(ms, ref width, ref height))
                {
                    return true;
                }
                else if (pictureType == ePictureType.WebP && IsWebP(ms, ref width, ref height))
                {
                    return true;
                }
                else if (pictureType == ePictureType.Ico && IsIcon(ms, ref width, ref height))
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        internal static byte[] ExtractImage(byte[] img, ref ePictureType type)
        {
            if (type == ePictureType.Emz ||
               type == ePictureType.Wmz)
            {
                try
                {
                    var ms = new MemoryStream(img);
                    var msOut = new MemoryStream();
                    const int bufferSize = 4096;
                    var buffer = new byte[bufferSize];
                    using (var z = new OfficeOpenXml.Packaging.Ionic.Zlib.GZipStream(ms, Packaging.Ionic.Zlib.CompressionMode.Decompress))
                    {
                        int size = 0;
                        do
                        {
                            size = z.Read(buffer, 0, bufferSize);
                            if (size > 0)
                            {
                                msOut.Write(buffer, 0, size);
                            }
                        }
                        while (size == bufferSize);
                        if (type == ePictureType.Emz) type = ePictureType.Emf;
                        else if (type == ePictureType.Wmz) type = ePictureType.Wmf;
                        return msOut.ToArray();
                    }
                }
                catch
                {
                    return img;
                }
            }
            return img;
        }

        private static bool IsIcon(MemoryStream ms, ref double width, ref double height)
        {
            using (var br = new BinaryReader(ms))
            {
                br.ReadInt16();
                var type = br.ReadInt16();
                if (type == 1)
                {
                    var imageCount = br.ReadInt16();
                    width = br.ReadByte();
                    height = br.ReadByte();
                    br.Close();
                    return true;
                }
                br.Close();
                return false;
            }
        }
        #region WebP
        private static bool IsWebP(MemoryStream ms, ref double width, ref double height)
        {
            using (var br = new BinaryReader(ms))
            {
                var riff=Encoding.ASCII.GetString(br.ReadBytes(4));
                var length = GetInt32BigEndian(br);
                var webP = Encoding.ASCII.GetString(br.ReadBytes(4));

                if (riff=="RIFF" && webP=="WEBP")
                {
                    var vp8= Encoding.ASCII.GetString(br.ReadBytes(4));
                    switch(vp8)
                    {
                        case "VP8 ":
                            var b = br.ReadBytes(10);
                            var w = br.ReadInt16();
                            width = w & 0x3FFF;
                            var hScale = w >> 14;
                            var h = br.ReadInt16();
                            height = h & 0x3FFF;
                            hScale = h >> 14;
                            break;
                        case "VP8X":
                            br.ReadBytes(8);
                            b = br.ReadBytes(6);
                            width = BitConverter.ToInt32(new byte[] { b[0], b[1], b[2], 0 }, 0) + 1;
                            height = BitConverter.ToInt32(new byte[] { b[3], b[4], b[5], 0 }, 0) + 1;
                            break;
                        case "VP8L":
                            br.ReadBytes(5);
                            b=br.ReadBytes(4);
                            width = (b[0] | (b[1] & 0x3F) << 8) + 1;
                            height = (b[1] >> 6 | b[2] << 2 | (b[3] & 0x0F) << 10) + 1;
                            break;
                    }
                }
            }
            return false;
        }
        #endregion
        #region Tiff
        private static bool IsTif(MemoryStream ms, ref double width, ref double height)
        {
            using (var br = new BinaryReader(ms))
            {
                var b = br.ReadBytes(2);
                var isBigEndian = Encoding.ASCII.GetString(b) == "MM";
                var identifier = GetTifInt16(br, isBigEndian);
                if (identifier == 42)
                {
                    var offset = GetTifInt32(br, isBigEndian);
                    ms.Position = offset;
                    var numberOfIdf = GetTifInt16(br, isBigEndian);
                    var ifds = new List<TifIfd>();
                    for (int i = 0; i < numberOfIdf; i++)
                    {
                        var ifd=new TifIfd()
                        {
                            Tag = GetTifInt16(br, isBigEndian),
                            Type = GetTifInt16(br, isBigEndian),
                            Count = GetTifInt32(br, isBigEndian),
                        };
                        if(ifd.Type==1 || ifd.Type==2 || ifd.Type == 6 || ifd.Type == 7)
                        {
                            ifd.ValueOffset = br.ReadByte();
                            br.ReadBytes(3);
                        }
                        else if (ifd.Type==3 || ifd.Type==8)
                        {
                            ifd.ValueOffset = GetTifInt16(br, isBigEndian);
                            br.ReadBytes(2);
                        }
                        else
                        {
                            ifd.ValueOffset = GetTifInt32(br, isBigEndian);
                        }
                        ifds.Add(ifd);
                    }

                    //int resolutionUnit=2;
                    //int XResolution = 1, YResolution = 1;
                    foreach (var ifd in ifds)
                    {
                        switch(ifd.Tag)
                        {
                            case 0x100:
                                width = ifd.ValueOffset;
                                break;
                            case 0x101:
                                height = ifd.ValueOffset;
                                break;
                            //case 0x128:
                            //    resolutionUnit= ifd.ValueOffset;
                            //    break;
                            //case 0x11A:
                            //    ms.Position=ifd.ValueOffset;
                            //    var l1 = GetTifInt32(br, isBigEndian);
                            //    var l2 = GetTifInt32(br, isBigEndian);
                            //    XResolution = l1;
                            //    break;
                            //case 0x11B:
                            //    ms.Position = ifd.ValueOffset;
                            //    l1 = GetTifInt32(br, isBigEndian);
                            //    l2 = GetTifInt32(br, isBigEndian);
                            //    YResolution = l1;
                            //    break;
                        }
                    }
                }
            }
            return width!=0 && height!=0;
        }
        private static short GetTifInt16(BinaryReader br, bool isBigEndian)
        {
            if(isBigEndian)
            {
                return GetInt16BigEndian(br);
            }
            else
            {
                return br.ReadInt16();
            }
        }
        private static int GetTifInt32(BinaryReader br, bool isBigEndian)
        {
            if (isBigEndian)
            {
                return GetInt32BigEndian(br);
            }
            else
            {
                return br.ReadInt32();
            }
        }
        #endregion
        #region Emf
        private static bool IsEmf(MemoryStream ms, ref double width, ref double height)
        {
            using (var br = new BinaryReader(ms))
            {
                if (br.PeekChar() == 1)
                {
                    var type = br.ReadInt32();
                    var length = br.ReadInt32();
                    var bounds = new int[4];
                    bounds[0] = br.ReadInt32();
                    bounds[1] = br.ReadInt32();
                    bounds[2] = br.ReadInt32();
                    bounds[3] = br.ReadInt32();
                    var frame = new int[4];
                    frame[0] = br.ReadInt32();
                    frame[1] = br.ReadInt32();
                    frame[2] = br.ReadInt32();
                    frame[3] = br.ReadInt32();

                    var signatureBytes = br.ReadBytes(4);
                    var signature = Encoding.ASCII.GetString(signatureBytes);
                    if (signature.Trim() == "EMF")
                    {
                        width = bounds[2] + 2;
                        height = bounds[3] + 2;
                        return true;
                    }
                }
            }
            return false;
        }
        #endregion
        #region Wmf
        private const double PIXELS_PER_TWIPS = 1D / 15D;
        private const double DEFAULT_TWIPS = 1440D;
        private static bool IsWmf(MemoryStream ms, ref double width, ref double height)
        {
            using (var br = new BinaryReader(ms))
            {
                var key = br.ReadUInt32();
                if (key == 0x9AC6CDD7)
                {
                    var HWmf = br.ReadInt16();
                    var bounds = new ushort[4];
                    bounds[0] = br.ReadUInt16();
                    bounds[1] = br.ReadUInt16();
                    bounds[2] = br.ReadUInt16();
                    bounds[3] = br.ReadUInt16();

                    var inch = br.ReadInt16();
                    width = bounds[2] - bounds[0];
                    height = bounds[3] - bounds[1];
                    if (inch != 0)
                    {
                        width *= (DEFAULT_TWIPS / inch) * PIXELS_PER_TWIPS;
                        height *= (DEFAULT_TWIPS / inch) * PIXELS_PER_TWIPS;
                    }
                    return width != 0 && height != 0;
                }
            }
            return false;
        }
        #endregion
        #region Png
        private static bool IsPng(MemoryStream ms, ref double width, ref double height)
        {
            using (var br = new BinaryReader(ms))
            {
                var signature = br.ReadBytes(8);
                if (signature.SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }))
                {
                    while (ms.Position < ms.Length)
                    {
                        var chunkType = ReadPngChunkHeader(br, out int length);
                        switch (chunkType)
                        {
                            case "IHDR":
                                width = GetInt32BigEndian(br);
                                height = GetInt32BigEndian(br);
                                br.ReadBytes(5); //Ignored bytes, Depth compression etc.
                                break;
                            case "pHYs":
                                float pixelsPerUnitX = GetInt32BigEndian(br);
                                float pixelsPerUnitY = GetInt32BigEndian(br);
                                var unitSpecifier = br.ReadByte();
                                if (unitSpecifier == 1)
                                {
                                    pixelsPerUnitX = pixelsPerUnitX / 39.36F / ExcelDrawing.STANDARD_DPI;
                                    pixelsPerUnitY = pixelsPerUnitY / 39.36F / ExcelDrawing.STANDARD_DPI;
                                }

                                width = width / pixelsPerUnitX;
                                height = height / pixelsPerUnitY;
                                br.Close();
                                return true;
                            default:
                                br.ReadBytes(length);
                                break;
                        }
                        var crc = br.ReadInt32();
                    }
                }

                br.Close();
            }
            return width!=0 && height!=0;
        }
        private static string ReadPngChunkHeader(BinaryReader br, out int length)
        {
            length = GetInt32BigEndian(br);
            var b = br.ReadBytes(4);
            var type = Encoding.ASCII.GetString(b);
            return type;
        }
        #endregion
        #region Svg
        private static bool IsSvg(MemoryStream ms, ref double width, ref double height)
        {
            try
            {
                using (var reader = new XmlTextReader(ms))
                {
                    while (reader.Read())
                    {
                        if (reader.LocalName == "svg" && reader.NodeType == XmlNodeType.Element)
                        {
                            var w = reader.GetAttribute("width");
                            var h = reader.GetAttribute("height");
                            var vb = reader.GetAttribute("viewBox");
                            reader.Close();
                            if (w == null || h == null)
                            {
                                if (vb == null)
                                {
                                    return false;
                                }
                                var bounds = vb.Split(new char[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
                                if (bounds.Length < 4)
                                {
                                    return false;
                                }
                                if (string.IsNullOrEmpty(w))
                                {
                                    w = bounds[2];
                                }
                                if (string.IsNullOrEmpty(h))
                                {
                                    h = bounds[3];
                                }
                            }
                            width = GetSvgUnit(w);
                            if (double.IsNaN(width)) return false;
                            height = GetSvgUnit(h);
                            if (double.IsNaN(height)) return false;
                            return true;
                        }
                    }
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private static double GetSvgUnit(string v)
        {
            var factor = 1D;
            if (v.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            {
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                factor = 1.25;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("pc", StringComparison.OrdinalIgnoreCase))
            {
                factor = 15;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
            {
                factor = 3.543307;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            {
                factor = 35.43307;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            {
                factor = 90;
                v = v.Substring(0, v.Length - 2);
            }
            if (double.TryParse(v, out double value))
            {
                return value * factor;
            }
            return double.NaN;
        }
        #endregion

        private static short GetInt16BigEndian(BinaryReader br)
        {
            var b = br.ReadBytes(2);
            return BitConverter.ToInt16(new byte[] { b[1], b[0] }, 0);
        }
        private static int GetInt32BigEndian(BinaryReader br)
        {
            var b = br.ReadBytes(4);
            return BitConverter.ToInt32(new byte[] { b[3], b[2], b[1], b[0] }, 0);
        }
    }
}
