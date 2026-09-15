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
using EPPlus.Export.Pdf.Helpers;
using System.IO;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfImageXObject : PdfObject
    {
        private readonly byte[] _bytes;
        internal int Width { get; }
        internal int Height { get; }
        internal string ColorSpace { get; }
        internal int BitsPerComponent { get; }
        internal string Filter { get; }
        internal string Decode { get; private set; }
        internal string DecodeParms { get; }

        internal bool HasSoftMask { get; }
        internal byte[] SoftMaskData { get; }
        internal int SoftMaskObjectNumber { get; set; }

        public PdfImageXObject(int objectNumber, byte[] imageBytes, int version = 0)
            : base(objectNumber, version)
        {
            byte[] icoDib = null;
            if (IcoDecoder.IsIco(imageBytes) && IcoDecoder.TryGetBestFrame(imageBytes, out byte[] icoFrame, out bool icoSelfContained))
            {
                if (icoSelfContained) imageBytes = icoFrame;
                else icoDib = icoFrame;
            }
            if (icoDib != null && IcoDecoder.TryDecodeDib(icoDib, out int icoW, out int icoH, out byte[] icoRgb, out byte[] icoAlpha))
            {
                Width = icoW;
                Height = icoH;
                BitsPerComponent = 8;
                ColorSpace = "/DeviceRGB";
                Filter = "FlateDecode";
                _bytes = PdfFlate.CompressLeaveOpen(icoRgb);
                if (icoAlpha != null)
                {
                    SoftMaskData = PdfFlate.CompressLeaveOpen(icoAlpha);
                    HasSoftMask = true;
                }
            }
            else if (JpegDecoder.IsJpeg(imageBytes))
            {
                _bytes = imageBytes;
                Filter = "DCTDecode";
                BitsPerComponent = 8;
                JpegDecoder.ReadJpegInfo(imageBytes, out int width, out int height, out int components, out bool adobe);
                Width = width;
                Height = height;
                if (components == 4)
                {
                    ColorSpace = "/DeviceCMYK";
                    if (adobe) Decode = "[ 1 0 1 0 1 0 1 0 ]";
                }
                else
                {
                    ColorSpace = components == 1 ? "/DeviceGray" : "/DeviceRGB";
                }
            }
            else if (PngDecoder.IsPng(imageBytes))
            {
                PngDecoder.ReadPngHeader(imageBytes, out int width, out int height, out int bitDepth, out int colorType, out int _);
                Width = width;
                Height = height;
                Filter = "FlateDecode";
                if (colorType == 6 || colorType == 4)
                {
                    BitsPerComponent = 8;
                    ColorSpace = colorType == 6 ? "/DeviceRGB" : "/DeviceGray";
                    PngDecoder.DecodePngWithAlpha(imageBytes, width, height, colorType, out byte[] color, out byte[] alpha);
                    _bytes = color;
                    SoftMaskData = alpha;
                    HasSoftMask = true;
                }
                else
                {
                    BitsPerComponent = bitDepth;
                    _bytes = PngDecoder.ReadPngIdat(imageBytes, out byte[] palette);
                    int colors;
                    switch (colorType)
                    {
                        case 0:
                            ColorSpace = "/DeviceGray";
                            colors = 1;
                            break;
                        case 3:
                            int hival = palette == null || palette.Length < 3 ? 0 : (palette.Length / 3) - 1;
                            ColorSpace = "[ /Indexed /DeviceRGB " + hival + " <" + PngDecoder.ToHex(palette) + "> ]";
                            colors = 1;
                            break;
                        default:
                            ColorSpace = "/DeviceRGB";
                            colors = 3;
                            break;
                    }
                    DecodeParms = "<< /Predictor 15 /Colors " + colors +
                                  " /BitsPerComponent " + bitDepth +
                                  " /Columns " + width + " >>";
                }
            }
            else if (BmpDecoder.TryDecode(imageBytes, out int bmpW, out int bmpH, out byte[] bmpRgb))
            {
                Width = bmpW;
                Height = bmpH;
                BitsPerComponent = 8;
                ColorSpace = "/DeviceRGB";
                Filter = "FlateDecode";
                _bytes = PdfFlate.CompressLeaveOpen(bmpRgb);
            }
            else if (GifDecoder.TryDecode(imageBytes, out int gifW, out int gifH, out byte[] gifRgb, out byte[] gifAlpha))
            {
                Width = gifW; Height = gifH;
                BitsPerComponent = 8;
                ColorSpace = "/DeviceRGB";
                Filter = "FlateDecode";
                _bytes = PdfFlate.CompressLeaveOpen(gifRgb);
                if (gifAlpha != null) { SoftMaskData = PdfFlate.CompressLeaveOpen(gifAlpha); HasSoftMask = true; }
            }
            else if (TiffDecoder.TryDecode(imageBytes, out int tifW, out int tifH, out byte[] tifRgb, out byte[] tifAlpha))
            {
                Width = tifW;
                Height = tifH;
                BitsPerComponent = 8;
                ColorSpace = "/DeviceRGB";
                Filter = "FlateDecode";
                _bytes = PdfFlate.CompressLeaveOpen(tifRgb);
                if (tifAlpha != null)
                {
                    SoftMaskData = PdfFlate.CompressLeaveOpen(tifAlpha);
                    HasSoftMask = true;
                }
            }
            else
            {
                _bytes = imageBytes;
                Filter = "DCTDecode";
                BitsPerComponent = 8;
                ColorSpace = "/DeviceRGB";
            }

        }

        private PdfImageXObject(int objectNumber, int version, byte[] deflatedGray, int width, int height)
            : base(objectNumber, version)
        {
            _bytes = deflatedGray;
            Width = width;
            Height = height;
            BitsPerComponent = 8;
            ColorSpace = "/DeviceGray";
            Filter = "FlateDecode";
        }

        internal static PdfImageXObject CreateSoftMask(int objectNumber, byte[] deflatedGray, int width, int height) => new PdfImageXObject(objectNumber, 0, deflatedGray, width, height);

        internal static bool CanEmbed(byte[] imageBytes)
        {
            if (IcoDecoder.IsIco(imageBytes))
            {
                if (!IcoDecoder.TryGetBestFrame(imageBytes, out byte[] frame, out bool selfContained))
                    return false;
                return selfContained ? CanEmbed(frame) : IcoDecoder.CanDecodeDib(frame);
            }
            if (JpegDecoder.IsJpeg(imageBytes))
            {
                return true;
            }
            if (PngDecoder.IsPng(imageBytes))
            {
                if (!PngDecoder.ReadPngHeader(imageBytes, out int _, out int _, out int bitDepth, out int colorType, out int interlace))
                    return false;
                if (interlace != 0) 
                    return false;
                if (colorType == 0 || colorType == 2 || colorType == 3) 
                    return true;
                if (colorType == 4 || colorType == 6) 
                    return bitDepth == 8;
                return false;
            }
            if (BmpDecoder.IsBmp(imageBytes))
            {
                return BmpDecoder.CanDecode(imageBytes);
            }
            if (GifDecoder.IsGif(imageBytes))
            {
                return GifDecoder.CanDecode(imageBytes);
            }
            if (TiffDecoder.IsTiff(imageBytes))
            {
                return TiffDecoder.CanDecode(imageBytes);
            }
            return false;
        }

        internal static bool ProducesSoftMask(byte[] imageBytes)
        {
            if (IcoDecoder.IsIco(imageBytes))
            {
                if (!IcoDecoder.TryGetBestFrame(imageBytes, out byte[] frame, out bool selfContained)) return false;
                return selfContained ? ProducesSoftMask(frame) : IcoDecoder.DibHasTransparency(frame);
            }
            if (PngDecoder.IsPng(imageBytes))
            {
                if (!PngDecoder.ReadPngHeader(imageBytes, out int _, out int _, out int bitDepth, out int colorType, out int interlace))
                    return false;
                if (interlace != 0) return false;
                return (colorType == 4 || colorType == 6) && bitDepth == 8;
            }
            if (GifDecoder.IsGif(imageBytes))
            {
                return GifDecoder.HasTransparency(imageBytes);
            }
            if (TiffDecoder.IsTiff(imageBytes))
            {
                return TiffDecoder.HasTransparency(imageBytes);
            }
            return false;
        }

        private string DictHeader()
        {
            string smask = HasSoftMask ? $" /SMask {SoftMaskObjectNumber.ToPdfStringF0()} 0 R" : "";
            string decode = string.IsNullOrEmpty(Decode) ? "" : $" /Decode {Decode}";
            string decodeParms = string.IsNullOrEmpty(DecodeParms) ? "" : $" /DecodeParms {DecodeParms}";
            return "<< /Type /XObject /Subtype /Image" +
                   $" /Width {Width.ToPdfStringF0()} /Height {Height.ToPdfStringF0()}" +
                   $" /ColorSpace {ColorSpace} /BitsPerComponent {BitsPerComponent.ToPdfStringF0()}" + smask + decode +
                   $" /Filter /{Filter}" + decodeParms +
                   $" /Length {_bytes.Length.ToPdfStringF0()} >>";
        }

        internal override string RenderDictionary()
        {
            return DictHeader() + $"\nstream\n<{_bytes.Length.ToPdfStringF0()} bytes of image data>\nendstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, DictHeader() + "\nstream\n");
            bw.Write(_bytes);
            WriteAscii(bw, "\nendstream");
        }
    }
}
