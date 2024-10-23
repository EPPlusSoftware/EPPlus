using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class BitmapHeader
    {
        enum BitMapType
        {
            BM,
            BA,
            CI,
            CP,
            IC,
            PT
        }
        internal enum CompressionMethod
        {
            BI_RGB = 0,
            BI_RLE8 = 1,
            BI_RLE4 = 2,
            BI_BITFIELDS = 3,
            BI_JPEG = 4,
            BI_PNG = 5,
            BI_ALPHABITFIELDS = 6,
            BI_CMYK = 7,
            BI_CMYKRLE8 = 8,
            BI_CMYKRLE4 = 9
        }

        //byte[] type = new byte[2];
        internal uint sizeOfHeader;
        internal int pixelWidth;
        internal int pixelHeight;
        internal ushort colorPlanes; //Must be 1?
        internal ushort colorDepth;
        internal uint compressionMethod;
        internal uint imageSize;
        internal int hRes; //Pixel per metre
        internal int vRes; //Pixel per metre
        internal uint nColors; // 0 defaults it to 2^n
        internal uint nImportantColors; //0 when all are. Generally ignored.

        internal CompressionMethod ReadCompression;
        internal byte[] ByteArrIfUnhandled = null;

        internal BitmapHeader(BinaryReader br, uint HeaderSize)
        {
            sizeOfHeader = br.ReadUInt32();

            if (sizeOfHeader == 40)
            {
                //Windows bitmapInfoHeader
                pixelWidth = br.ReadInt32();
                pixelHeight = br.ReadInt32();
                colorPlanes = br.ReadUInt16();
                colorDepth = br.ReadUInt16();
                compressionMethod = br.ReadUInt32();
                ReadCompression = (CompressionMethod)compressionMethod;

                imageSize = br.ReadUInt32();
                hRes = br.ReadInt32();
                vRes = br.ReadInt32();
                nColors = br.ReadUInt32();
                nImportantColors = br.ReadUInt32();
            }
            else
            {
                ByteArrIfUnhandled = br.ReadBytes((int)HeaderSize - 4);
            }
        }

        internal int offset;

        //Reading raw file
        internal BitmapHeader(BinaryReader br)
        {
            var sign = Encoding.ASCII.GetString(br.ReadBytes(2));    //BM for a Windows bitmap

            var size = br.ReadInt32();
            var reserved = br.ReadBytes(4);
            offset = br.ReadInt32();

            //Windows bitmapInfoHeader
            sizeOfHeader = br.ReadUInt32();

            if (sizeOfHeader == 40)
            {
                pixelWidth = br.ReadInt32();
                pixelHeight = br.ReadInt32();
                colorPlanes = br.ReadUInt16();
                colorDepth = br.ReadUInt16();
                compressionMethod = br.ReadUInt32();
                ReadCompression = (CompressionMethod)compressionMethod;

                imageSize = br.ReadUInt32();
                hRes = br.ReadInt32();
                vRes = br.ReadInt32();
                nColors = br.ReadUInt32();
                nImportantColors = br.ReadUInt32();
            }
            else
            {
                ByteArrIfUnhandled = br.ReadBytes((int)sizeOfHeader);
            }
        }


        internal void WriteBytes(BinaryWriter bw)
        {
            if(ByteArrIfUnhandled == null)
            {
                bw.Write(sizeOfHeader);
                bw.Write(pixelWidth);
                bw.Write(pixelHeight);
                bw.Write(colorPlanes);
                bw.Write(colorDepth);
                bw.Write((uint)ReadCompression);
                bw.Write(imageSize);
                bw.Write(hRes);
                bw.Write(vRes);
                bw.Write(nColors);
                bw.Write(nImportantColors);
            }
            else
            {
                bw.Write(ByteArrIfUnhandled);
            }
        }
    }
}
