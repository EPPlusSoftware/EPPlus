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
            BI_RGB,
            BI_RLE8,
            BI_RLE4,
            BI_BITFIELDS,
            BI_JPEG,
            BI_PNG,
            BI_ALPHABITFIELDS,
            BI_CMYK,
            BI_CMYKRLE8,
            BI_CMYKRLE4
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
        internal byte[] ByteArrIfUnhandled;

        internal BitmapHeader(BinaryReader br, uint HeaderSize)
        {
            if(HeaderSize == 40)
            {
                //Windows bitmapInfoHeader
                sizeOfHeader = br.ReadUInt32();
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
                ByteArrIfUnhandled = br.ReadBytes((int)HeaderSize);
            }

            //br.Read(type, 0, 2);

            //var strType = Encoding.ASCII.GetString(type);
            //sizeOfFile = br.ReadUInt32();

        }
    }
}
