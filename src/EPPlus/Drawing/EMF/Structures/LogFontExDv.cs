using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class LogFontExDv : LogFontEx
    {
        internal DesignVector dv;

        internal LogFontExDv() : base()
        {
            dv = new DesignVector();
        }
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
