/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_STRETCHDIBITS : EMR_RECORD
    {
        internal RectLObject Bounds;
        internal int xDest;
        internal int yDest;
        internal int xSrc;
        internal int ySrc;
        internal int cxSrc;
        internal int cySrc;
        internal uint offBmiSrc;
        internal uint cbBmiSrc;
        internal uint offBitsSrc;
        internal uint cbBitsSrc;
        internal uint UsageSrc;
        internal uint InternalBltRasterOperation;
        internal int cxDest;
        internal int cyDest;
        internal byte[] BmiSrc;
        internal byte[] _bitsSrc;

        //Variables for replacing/resizing the image.
        //If behaves strangely reset SETWORLDTRANSFORM and MODIFYWORLDTRANSFORM records in EmfImage file.
        internal double MaxHeight = 75;
        internal double MaxWidth = 111;

        internal BitmapHandler ExtractedBmp;
        
        internal byte[] BitsSrc
        {
            get
            {
                return _bitsSrc;
            }
            set
            {
                Size -= (uint)_bitsSrc.Length;
                _bitsSrc = value;
                cbBitsSrc = (uint)_bitsSrc.Length;
                Size += (uint)_bitsSrc.Length;
                if(Size % 4 != 0)
                {
                    int paddingBytes = (int)(4 - (Size % 4)) % 4;
                    EndPadding = new byte[paddingBytes];
                    Size += (uint)paddingBytes;
                }
            }
        }
        internal BitmapInformationHeader bitMapHeader;

        internal byte[] Padding1;

        byte[] _padding2 = new byte[0];

        internal byte[] Padding2
        {
            get
            {
                return _padding2;
            }
            set
            {
                Size -= (uint)_padding2.Length;
                offBitsSrc -= (uint)_padding2.Length;
                _padding2 = value;
                Size += (uint)_padding2.Length;

                if (Size % 4 != 0)
                {
                    int paddingBytes = (int)(4 - (Size % 4)) % 4;
                    EndPadding = new byte[paddingBytes];
                    Size += (uint)paddingBytes;
                }
                offBitsSrc += (uint)_padding2.Length;
            }
        }
        internal byte[] EndPadding;

        internal EMR_STRETCHDIBITS(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            var startOfRecord = br.BaseStream.Position - 8;

            Bounds = new RectLObject(br);
            xDest = br.ReadInt32();
            yDest = br.ReadInt32();
            xSrc = br.ReadInt32();
            ySrc = br.ReadInt32();
            cxSrc = br.ReadInt32();
            cySrc = br.ReadInt32();
            offBmiSrc = br.ReadUInt32();
            cbBmiSrc = br.ReadUInt32();
            offBitsSrc = br.ReadUInt32();
            cbBitsSrc = br.ReadUInt32();
            UsageSrc = br.ReadUInt32();
            InternalBltRasterOperation = br.ReadUInt32();
            cxDest = br.ReadInt32();
            cyDest = br.ReadInt32();

            //There's undefined variable space here, ensure we reach the header
            var startOfHeader = startOfRecord + offBmiSrc;
            int padding1Length = 0;
            if(br.BaseStream.Position < startOfHeader)
            {
                padding1Length = (int)(startOfHeader - br.BaseStream.Position);
                Padding1 = new byte[padding1Length];
                br.Read(Padding1, 0, padding1Length);
            }

            bitMapHeader = new BitmapInformationHeader(br, cbBmiSrc);

            //There's undefined variable space here, ensure we reach the bitmapSpace
            var startOfBitmapBits = startOfRecord + offBitsSrc;
            int padding2Length = 0;

            if (br.BaseStream.Position < startOfBitmapBits)
            {
                padding2Length = (int)(startOfBitmapBits - br.BaseStream.Position);
                _padding2 = new byte[padding2Length];
                br.Read(_padding2, 0, padding2Length);
            }

            //Source bitmap bits
            _bitsSrc = br.ReadBytes((int)cbBitsSrc);

            //Helper property to Order information as a valid bitmap outside emf
            ExtractedBmp = new BitmapHandler();
            ExtractedBmp.fileHeader = new BitMapFileHeader();
            ExtractedBmp.informationHeader = bitMapHeader;
            ExtractedBmp.OptionalData = Padding2;
            ExtractedBmp.PixelArray = _bitsSrc;
            ExtractedBmp.fileHeader.Offset = (int)(14/*FileHeader*/+ bitMapHeader.sizeOfHeader + padding2Length);
            ExtractedBmp.fileHeader.Size = (int)(bitMapHeader.sizeOfHeader + 14/*FileHeader*/ + cbBitsSrc + padding2Length);

            int tempPadding = (int)((position + Size) - br.BaseStream.Position);
            if (tempPadding < 0)
            {
                EndPadding = new byte[0];
                return;
            }
            EndPadding = br.ReadBytes(tempPadding);
        }

        internal void UpdateToImage(string fileName)
        {
            var img = new ExcelImage(fileName);

            switch (img.Type.Value)
            {
                case ePictureType.Bmp:
                    ReadBmpAndUpdateImage(File.ReadAllBytes(fileName));
                    break;
                default:
                    {
                        throw new NotSupportedException($"{fileName} could not be read. The filetype: {img.Type} is not supported in digital signatures. please use a .bmp file.");
                    }
            }
        }

        private void UpdateImage(BitmapHandler handler, bool centerImage = false, bool adjustYOriginToHeight = true)
        {
            bitMapHeader = handler.informationHeader;
            cbBmiSrc = bitMapHeader.sizeOfHeader;
            Padding2 = handler.OptionalData;
            BitsSrc = handler.PixelArray;

            cxSrc = handler.informationHeader.pixelWidth;
            cySrc = handler.informationHeader.pixelHeight;
            RecalculateImage(centerImage, adjustYOriginToHeight);
        }

        /// <summary>
        /// Please Note: Assumes world origin has been reset.
        /// </summary>
        internal void RecalculateImage(bool centerImage = false, bool adjustYOriginToHeight = true)
        {
            var xSource = cxSrc;
            var ySource = cySrc;

            double xRatio = MaxWidth / (double)xSource;
            double yRatio = MaxHeight / (double)ySource;

            double ratio = xRatio < yRatio ? xRatio : yRatio;

            cxDest = Convert.ToInt32(xSource * ratio);
            cyDest = Convert.ToInt32(ySource * ratio);

            //minor pixel adjustment to be closer to excel
            if(cyDest == MaxHeight)
            {
                if(cxDest == cyDest)
                {
                    cxDest -= 1;
                }
                cyDest -= 1;
            }

            if(centerImage)
            {
                xDest = Convert.ToInt32((MaxWidth - (xSource * ratio)) / 2);
                yDest = Convert.ToInt32((MaxHeight - (ySource * ratio)) / 2);
            }
            else if(adjustYOriginToHeight)
            {
                if(cyDest <= MaxHeight)
                {
                    yDest = Convert.ToInt32(MaxHeight - (ySource * ratio));
                }
            }
        }

        internal void ReadBmpAndUpdateImage(byte[] bmp, bool centerImage = false, bool adjustYOriginToHeight = true)
        {
            ExtractedBmp = new BitmapHandler(bmp);
            UpdateImage(ExtractedBmp, centerImage, adjustYOriginToHeight);
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            Bounds.WriteBytes(bw);
            bw.Write(xDest);
            bw.Write(yDest);
            bw.Write(xSrc);
            bw.Write(ySrc);
            bw.Write(cxSrc);
            bw.Write(cySrc);
            bw.Write(offBmiSrc);
            bw.Write(cbBmiSrc);
            bw.Write(offBitsSrc);
            bw.Write(cbBitsSrc);
            bw.Write(UsageSrc); 
            bw.Write(InternalBltRasterOperation);
            bw.Write(cxDest);
            bw.Write(cyDest);
            if(Padding1 != null)
            {
                bw.Write(Padding1);
            }
            bitMapHeader.WriteBytes(bw);
            if (Padding2 != null)
            {
                bw.Write(Padding2);
            }
            bw.Write(BitsSrc);
            bw.Write(EndPadding);
        }
    }
}
