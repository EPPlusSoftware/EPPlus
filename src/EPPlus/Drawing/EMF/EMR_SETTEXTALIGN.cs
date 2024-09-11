using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_SETTEXTALIGN : EMR_RECORD
    {
        internal byte[] TextAlignmentMode;

        public EMR_SETTEXTALIGN(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            TextAlignmentMode = br.ReadBytes(4);
        }
    }
}
