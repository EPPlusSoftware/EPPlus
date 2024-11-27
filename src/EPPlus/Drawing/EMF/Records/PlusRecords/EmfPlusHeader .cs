using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.EMF.PlusStructure;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

namespace OfficeOpenXml.Drawing.EMF.Records
{
    internal class EmfPlusHeader : EmfPlusRecord
    {
        //If true the file contains two sets of records. Both SHOULD completely define the graphics.
        //If false graphics are defined by EMF+ records alone.
        internal bool IsDual = false;
        internal EmfPlusGraphicsVersionObject GraphicsVersion;
        internal byte[] EmfPlusFlags;

        internal uint LogicalDpiX;
        internal uint LogicalDpiY;

        internal EmfPlusHeader(BinaryReader br) : base(br, RECORD_TYPES_PLUS.EmfPlusHeader)
        {
            //Last becomes first because of little-endian

            //Last bit in plusflags defines if the metafile is EMF+ Dual
            IsDual = (PlusFlags[0] & (1 << 1 - 1)) != 0;
            GraphicsVersion = new EmfPlusGraphicsVersionObject(br);

            //Last bit defines device as video if set, printer if not.
            EmfPlusFlags = br.ReadBytes(4);

            LogicalDpiX = br.ReadUInt32();
            LogicalDpiY = br.ReadUInt32();
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(GraphicsVersion.bytes);
            bw.Write(EmfPlusFlags);
            bw.Write(LogicalDpiX);
            bw.Write(LogicalDpiY);
        }
    }
}
