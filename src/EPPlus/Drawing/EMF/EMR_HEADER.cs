using System.IO;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_HEADER : EMR_RECORD
    {
        internal byte[] Bounds;             //16
        internal byte[] Frame;              //16
        internal byte[] RecordSignature;    //4
        internal byte[] Version;            //4
        internal uint   Bytes;              //4         //Filesize
        internal uint   Records;            //4         //List Size
        internal ushort Handles;            //2         //number of graphics objects
        internal byte[] Reserved;           //2
        internal byte[] nDescription;       //4
        internal byte[] offDescription;     //4
        internal uint   nPalEntries;        //4         //Found in EOF
        internal byte[] Device;             //8
        internal byte[] Millimeters;        //8
        internal byte[] cbPixelFormat;      //4
        internal byte[] offPixelFormat;     //4
        internal byte[] bOpenGL;            //4
        internal byte[] MicroMetersX;       //4
        internal byte[] MicroMetersY;       //4

        public EMR_HEADER(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            
            Bounds = br.ReadBytes(16);
            Frame = br.ReadBytes(16);
            RecordSignature = br.ReadBytes(4);
            Version = br.ReadBytes(4);
            Bytes = br.ReadUInt32();
            Records = br.ReadUInt32();
            Handles = br.ReadUInt16();
            Reserved = br.ReadBytes(2);
            nDescription = br.ReadBytes(4);
            offDescription = br.ReadBytes(4);
            nPalEntries = br.ReadUInt32();
            Device = br.ReadBytes(8);
            Millimeters = br.ReadBytes(8);
            cbPixelFormat = br.ReadBytes(4);
            offPixelFormat = br.ReadBytes(4);
            bOpenGL = br.ReadBytes(4);
            MicroMetersX = br.ReadBytes(4);
            MicroMetersY = br.ReadBytes(4);
        }

        public EMR_HEADER(List<EMR_RECORD> Records)
        {
            Type = RECORD_TYPES.EMR_HEADER;
            Bounds = new byte[16] { 0x13, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x4b, 0x00, 0x00, 0x00, 0x30, 0x00, 0x00, 0x00 };
            Frame = new byte[16] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xBD, 0x08, 0x00, 0x00, 0x90, 0x06, 0x00, 0x00 };
            RecordSignature = new byte[4] { 0x20, 0x45, 0x4D, 0x46 };
            Version = new byte[4] { 0x00, 0x00, 0x01, 0x00 };
            Reserved = new byte[2] { 0x00, 0x00 };
            nDescription = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            offDescription = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            Device = new byte[8] { 0x00, 0x14, 0x00, 0x00, 0xA0, 0x05, 0x00, 0x00 };
            Millimeters = new byte[8] { 0xA9, 0x04, 0x00, 0x00, 0x50, 0x01, 0x00, 0x00 };
            cbPixelFormat = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            offPixelFormat = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
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
                //switch (record.Type)
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

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(Bounds);
            bw.Write(Frame);
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
        }

    }
}
