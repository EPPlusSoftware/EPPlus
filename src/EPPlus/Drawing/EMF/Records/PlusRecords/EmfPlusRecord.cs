using OfficeOpenXml.Encryption;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF.Records
{
    internal class EmfPlusRecord
    {
        internal RECORD_TYPES_PLUS Type; //2
        internal byte[] PlusFlags;       //2
        /// <summary>
        /// Defines size of entire record
        /// </summary>
        internal uint Size;              //4
        /// <summary>
        /// Size without the Invariant part of given record
        /// </summary>
        internal uint DataSize;          //4
        internal byte[] data;       //This byte array is used for records not yet implemented to preserve data.
        internal long position = 0;

        internal EmfPlusRecord() { }

        internal EmfPlusRecord(BinaryReader br, RECORD_TYPES_PLUS type, bool readData = false)
        {
            position = br.BaseStream.Position - 4;
            Type = type;
            PlusFlags = br.ReadBytes(2);
            Size = br.ReadUInt32();
            DataSize = br.ReadUInt32();

            if (readData && Size > 12)
            {
                data = br.ReadBytes((int)Size - 12);
            }
        }

        internal virtual void WriteBytes(BinaryWriter bw)
        {
            bw.Write((ushort)Type);
            bw.Write(PlusFlags);
            bw.Write(Size);
            bw.Write(DataSize);

            if (data != null)
            {
                bw.Write(data);
            }
        }
    }
}
