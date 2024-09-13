using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class LogFontExDv : LogFontEx
    {
        DesignVector dv;

        internal LogFontExDv(BinaryReader br): base(br)
        {
            dv = new DesignVector(br);
        }
        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            dv.WriteBytes(bw);
        }
    }
}
