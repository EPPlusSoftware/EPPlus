using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_SETMAPMODE : EMR_RECORD
    {
        internal MapMode MapMode;

        internal EMR_SETMAPMODE(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            MapMode = (MapMode)br.ReadUInt32();
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write((uint)MapMode);
        }
    }
}
