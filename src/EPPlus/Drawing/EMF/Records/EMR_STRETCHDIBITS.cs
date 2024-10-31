using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.IO;
using System.Text;

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
        private int _targetHeight = 75;
        private int _targetWidth = 111;

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
            if(br.BaseStream.Position < startOfHeader)
            {
                int padding = (int)(startOfHeader - br.BaseStream.Position);
                Padding1 = new byte[padding];
                br.Read(Padding1, 0, padding);
            }

            //Should not be neccesary
            br.BaseStream.Position = startOfHeader;

            bitMapHeader = new BitmapInformationHeader(br, cbBmiSrc);

            //There's undefined variable space here, ensure we reach the bitmapSpace
            var startOfBitmapBits = startOfRecord + offBitsSrc;
            if (br.BaseStream.Position < startOfBitmapBits)
            {
                int padding = (int)(startOfBitmapBits - br.BaseStream.Position);
                _padding2 = new byte[padding];
                br.Read(_padding2, 0, padding);
            }

            //Should not be neccesary
            br.BaseStream.Position = startOfBitmapBits;

            //Source bitmap bits
            _bitsSrc = br.ReadBytes((int)cbBitsSrc);

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

        private void UpdateImage(BitmapHandler handler)
        {
            bitMapHeader = handler.informationHeader;
            cbBmiSrc = bitMapHeader.sizeOfHeader;
            Padding2 = handler.OptionalData;
            BitsSrc = handler.PixelArray;

            cxSrc = handler.informationHeader.pixelWidth;
            cySrc = handler.informationHeader.pixelHeight;
            RecalculateImage();
        }

        internal void RecalculateImage()
        {
            var xSource = cxSrc;
            var ySource = cySrc;

            double xRatio = (double)111 / (double)xSource;
            double yRatio = (double)75 / (double)ySource;

            double ratio = xRatio < yRatio ? xRatio : yRatio;

            cxDest = Convert.ToInt32(xSource * ratio);
            cyDest = Convert.ToInt32(ySource * ratio);

            xDest = Convert.ToInt32((111f - (xSource * ratio)) / 2);
            yDest = Convert.ToInt32((75f - (ySource * ratio)) / 2);
        }

        internal void ReadBmpAndUpdateImage(byte[] bmp)
        {
            var handler = new BitmapHandler(bmp);
            UpdateImage(handler);
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
