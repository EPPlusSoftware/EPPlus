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
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml.Packaging.Ionic.Zip;
using System.Text;
using System;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.Utils.FileUtils;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_HEADER : EMR_RECORD
    {
        /// <summary>
        ///  Bounds in logical units (LU) (LU is equivalent to pixels if MapMode defined by EMR_SETMAPMODE is MM_Text.)
        /// </summary>
        internal RectLObject Bounds;             //16
        /// <summary>
        /// Frame in 0.1 mm units
        /// </summary>
        internal RectLObject Frame;              //16
        internal byte[] RecordSignature;    //4
        internal byte[] Version;            //4
        internal uint   Bytes;              //4         //Filesize
        internal uint   Records;            //4         //List FontSize
        internal ushort Handles;            //2         //number of graphics objects
        internal byte[] Reserved;           //2
        internal uint nDescription;       //4
        internal uint offDescription;     //4
        internal uint   nPalEntries;        //4         //Found in EOF
        internal byte[] Device;             //8
        internal byte[] Millimeters;        //8
        internal uint cbPixelFormat;      //4
        internal uint offPixelFormat;     //4
        internal byte[] bOpenGL;            //4
        internal byte[] MicroMetersX;       //4
        internal byte[] MicroMetersY;       //4

        internal string DescriptionString;
        internal void SetDescriptionString(string text)
        {
            text = text + "\0";
            DescriptionString = text;
            offDescription = headerSize;
            nDescription = (uint)DescriptionString.Length;
            Size += nDescription * 2;
        }

        internal byte[] PixelFormatDescriptor;

        internal string headerType = "Emf_MetafileHeader";
        internal uint headerSize;

        internal float inchesX;
        internal float inchesY;
        internal float Ppi;

        internal double MilimetersPerPixelX;
        internal double MilimetersPerPixelY;

        internal EMR_HEADER(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            if (Size >= 84)
            {
                headerSize = Size;

                Bounds = new RectLObject(br);
                Frame = new RectLObject(br);
                RecordSignature = br.ReadBytes(4);
                Version = br.ReadBytes(4);
                Bytes = br.ReadUInt32();
                Records = br.ReadUInt32();
                Handles = br.ReadUInt16();
                Reserved = br.ReadBytes(2);
                nDescription = br.ReadUInt32();
                offDescription = br.ReadUInt32();
                nPalEntries = br.ReadUInt32();
                Device = br.ReadBytes(8);
                Millimeters = br.ReadBytes(8);

                //Valid description?
                if (offDescription >= 88 && offDescription + (nDescription * 2) <= Size)
                {
                    headerSize = offDescription;
                }

                if (headerSize >= 100)
                {
                    //Header is SomeKind of headerExtension
                    cbPixelFormat = br.ReadUInt32();
                    offPixelFormat = br.ReadUInt32();
                    bOpenGL = br.ReadBytes(4);

                    if (offPixelFormat >= 100 && offPixelFormat + cbPixelFormat <= Size)
                    {
                        if(offPixelFormat < headerSize)
                        {
                            headerSize = offPixelFormat;
                        }
                    }

                    if(headerSize >= 108)
                    {
                        headerType += "Extension2";
                        //TODO: Define how to determine extension2
                        MicroMetersX = br.ReadBytes(4);
                        MicroMetersY = br.ReadBytes(4);
                    }
                    else
                    {
                        headerType += "Extension1";
                    }
                }

                if(Size != headerSize)
                {
                    var pos = br.BaseStream.Position;

                    if (offDescription != pos)
                    {
                        br.BaseStream.Position = offDescription;
                    }
                    DescriptionString = BinaryHelper.GetString(br, (nDescription * 2), Encoding.Unicode);
                    if(offPixelFormat != 0)
                    {
                        br.BaseStream.Position = offPixelFormat;
                        PixelFormatDescriptor = br.ReadBytes((int)cbPixelFormat);
                    }

                    if(br.BaseStream.Position != Size)
                    {
                        //Something weird, likely EMF+ record
                        br.BaseStream.Position = Size;
                    }
                }

                var cxMili = BitConverter.ToUInt32(Millimeters, 0);
                var cyMili = BitConverter.ToUInt32(Millimeters, 4);

                //Convert milimeters to inches
                inchesX = cxMili * 0.0393700787f;
                inchesY = cyMili * 0.0393700787f;

                var cx = BitConverter.ToUInt32(Device, 0);
                var cy = BitConverter.ToUInt32(Device, 4);

                Ppi = cx / inchesX;
                MilimetersPerPixelX = cxMili / (double)cx;
                MilimetersPerPixelY = cyMili / (double)cy;
            }
            else
            {
                throw new BadReadException("Emf-Header MUST be larger than or equal to 84");
            }
        }

        internal EMR_HEADER(List<EMR_RECORD> Records)
        {
            Type = RECORD_TYPES.EMR_HEADER;
            Bounds = new RectLObject(13, 2, 75, 30);
            Frame = new RectLObject(0, 0, 2237, 1680);
            RecordSignature = new byte[4] { 0x20, 0x45, 0x4D, 0x46 };
            Version = new byte[4] { 0x00, 0x00, 0x01, 0x00 };
            Reserved = new byte[2] { 0x00, 0x00 };
            nDescription =  0;
            offDescription = 0;
            Device = new byte[8] { 0x00, 0x14, 0x00, 0x00, 0xA0, 0x05, 0x00, 0x00 };
            Millimeters = new byte[8] { 0xA9, 0x04, 0x00, 0x00, 0x50, 0x01, 0x00, 0x00 };
            cbPixelFormat = 0;
            offPixelFormat = 0;
            bOpenGL = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            MicroMetersX = new byte[4] { 0x28, 0x34, 0x12, 0x00 };
            MicroMetersY = new byte[4] { 0x80, 0x20, 0x05, 0x00 };
            Size = 4 + 4 + 16 + 16 + 4 + 4 + 4 + 4 + 2 + 2 + 4 + 4 + 4 + 8 + 8 + 4 + 4 + 4 + 4 + 4;
            this.Records = (uint)Records.Count;
            var eof = Records[Records.Count - 1] as EMR_EOF;
            nPalEntries = eof.nPalEntries;
            Bytes = 0;
            Handles = 3;
            foreach (var record in Records)
            {
                //switch (record.BulletType)
                //{
                //    case RECORD_TYPES.EMR_CREATEPEN:
                //    case RECORD_TYPES.EMR_EXTCREATEPEN:
                //    case RECORD_TYPES.EMR_CREATEBRUSHINDIRECT:
                //    case RECORD_TYPES.EMR_CREATEDIBPATTERNBRUSHPT:
                //    case RECORD_TYPES.EMR_CREATEMONOBRUSH:
                //    case RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW:
                //    case RECORD_TYPES.EMR_CREATEPALETTE:
                //    case RECORD_TYPES.EMR_STRETCHDIBITS:
                //    case RECORD_TYPES.EMR_STRETCHBLT:
                //    case RECORD_TYPES.EMR_CREATECOLORSPACE:
                //    case RECORD_TYPES.EMR_CREATECOLORSPACEW:
                //        Handles++;
                //        break;
                //    case RECORD_TYPES.EMR_DELETECOLORSPACE:
                //    case RECORD_TYPES.EMR_DELETEOBJECT:
                //        Handles--;
                //        break;
                //}
                Bytes += record.Size;
            }
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            Bounds.WriteBytes(bw);
            Frame.WriteBytes(bw);
            bw.Write(RecordSignature);
            bw.Write(Version);
            bw.Write(Bytes);
            bw.Write(Records);
            bw.Write(Handles);
            bw.Write(Reserved);
            bw.Write(nDescription);
            bw.Write(offDescription);
            bw.Write(nPalEntries);
            bw.Write(Device);
            bw.Write(Millimeters);
            bw.Write(cbPixelFormat);
            bw.Write(offPixelFormat);
            bw.Write(bOpenGL);
            bw.Write(MicroMetersX);
            bw.Write(MicroMetersY);
            if(string.IsNullOrEmpty(DescriptionString) == false)
            {
                bw.Write(BinaryHelper.GetByteArray(DescriptionString, Encoding.Unicode));
            }
        }

    }
}
