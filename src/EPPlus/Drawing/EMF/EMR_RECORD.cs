using System.IO;
using System.Diagnostics;

namespace OfficeOpenXml.Drawing.EMF
{
    [DebuggerDisplay("Type: {Type}, Size: {Size}")]
    internal class EMR_RECORD
    {
        internal RECORD_TYPES Type; //4
        internal uint Size;         //4
        internal byte[] data;       //Variable
        internal long position = 0;

        public EMR_RECORD() { }

        public EMR_RECORD(BinaryReader br, uint TypeValue, bool readData = false)
        {
            position = br.BaseStream.Position - 4;
            Type = (RECORD_TYPES)TypeValue;
            Size = br.ReadUInt32();
            if (readData && Size > 8)
            {
                data = br.ReadBytes((int)Size - 8);
            }
        }

        public virtual void WriteBytes(BinaryWriter bw)
        {
            bw.Write((uint)Type);
            bw.Write(Size);
            if (data != null)
            {
                bw.Write(data);
            }
        }

    }
}
