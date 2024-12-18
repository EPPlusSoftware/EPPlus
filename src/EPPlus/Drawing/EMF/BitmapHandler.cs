using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class BitmapHandler
    {
        internal BitMapFileHeader fileHeader;
        internal BitmapInformationHeader informationHeader;

        //ExtraBitMasks + ColorTable + GAP1 variable size
        internal byte[] OptionalData = new byte[0];
        /// <summary>
        /// Gap2 + ICC color profile if they exist
        /// </summary>
        internal byte[] OptionalData2 = new byte[0];

        internal byte[] PixelArray;

        internal BitmapHandler()
        {
        }

        internal BitmapHandler(byte[] fileBytes)
        {
            ReadBitmap(fileBytes);
        }

        internal void ReadBitmap(byte[] fileBytes)
        {
            using (var ms = new MemoryStream(fileBytes))
            {
                using (var br = new BinaryReader(ms))
                {
                    br.BaseStream.Seek(0, SeekOrigin.Begin);

                    fileHeader = new BitMapFileHeader(br);
                    informationHeader = new BitmapInformationHeader(br);

                    //Could be broken out to individual data. ExtraBitMasks only for BI_BITFIELDS/BI_ALPHABITFIELDS etc.
                    var sizeOfOptionalData = fileHeader.Offset - (int)br.BaseStream.Position;
                    OptionalData = br.ReadBytes(sizeOfOptionalData);

                    int pixelArrLen = (int)informationHeader.imageSize;
                    if (pixelArrLen == 0 && informationHeader.ReadCompression == BitmapInformationHeader.CompressionMethod.BI_RGB)
                    {
                        pixelArrLen = (int)fileBytes.Length - (int)br.BaseStream.Position;
                    }
                    PixelArray = br.ReadBytes(pixelArrLen);
                   
                    int remainingData = fileHeader.Size - (int)br.BaseStream.Position;
                    OptionalData2 = br.ReadBytes(remainingData);
                }
            }
        }

        internal byte[] GetBitMapBytes()
        {
            var bitArr = new byte[fileHeader.Size];
            using (var ms = new MemoryStream(bitArr))
            {
                var bw = new BinaryWriter(ms);
                fileHeader.WriteBytes(bw);
                informationHeader.WriteBytes(bw);
                bw.Write(OptionalData);
                bw.Write(PixelArray);
                bw.Write(OptionalData2);
            }
            return bitArr;
        }
    }
}
