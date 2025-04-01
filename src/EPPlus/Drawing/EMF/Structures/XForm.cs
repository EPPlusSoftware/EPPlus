using System.IO;

namespace OfficeOpenXml.Drawing.EMF.Structures
{
    //Used to transform world space to page space etc.

    //X' = M11 * X + M21 * Y + Dx
    //Y' = M12 * X + M22 * Y + Dy

    internal class XForm
    {
        internal float M11;
        internal float M12;
        internal float M21;
        internal float M22;
        internal float Dx;
        internal float Dy;

        internal XForm(BinaryReader br)
        {
            M11 = br.ReadSingle();
            M12 = br.ReadSingle();
            M21 = br.ReadSingle();
            M22 = br.ReadSingle();
            Dx = br.ReadSingle();
            Dy = br.ReadSingle();
        }

        internal void WriteBytes(BinaryWriter bw)
        {
            bw.Write(M11);
            bw.Write(M12);
            bw.Write(M21);
            bw.Write(M22);
            bw.Write(Dx);
            bw.Write(Dy);
        }
    }
}
