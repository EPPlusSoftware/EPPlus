using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class BitMapFileHeader
    {
        /// <summary>
        /// All but BM is OS/2 operating system.
        /// </summary>
        internal enum BitMapType : UInt16
        {
            /// <summary>
            /// Windows bitmap.
            /// </summary>
            BM = 0x4d42,
            /// <summary>
            /// Bitmap Array
            /// </summary>
            BA = 0x4d41,
            /// <summary>
            /// Color Icon
            /// </summary>
            CI = 0x4349,
            /// <summary>
            /// Color Pointer
            /// </summary>
            CP = 0x4350,
            /// <summary>
            /// Icon
            /// </summary>
            IC = 0x4943,
            /// <summary>
            /// Pointer
            /// </summary>
            PT = 0x5054
        }

        /// <summary>
        /// Type of bitmap
        /// </summary>
        internal BitMapType Signature;
        /// <summary>
        /// Byte size of bitmap
        /// </summary>
        internal int Size;
        /// <summary>
        /// Application dependent reserved space
        /// </summary>
        internal byte[] Reserved1 = new byte[2];
        internal byte[] Reserved2 = new byte[2];
        /// <summary>
        /// Offse to pixel array/image data
        /// </summary>
        internal int Offset;

        internal BitMapFileHeader()
        {
            Signature = BitMapType.BM;   
        }

        internal BitMapFileHeader(MemoryStream ms)
        {
            using(var br = new BinaryReader(ms))
            {
                if(IsBmp(br))
                {
                    Size = br.ReadInt32();
                    Reserved1 = br.ReadBytes(2);
                    Reserved2 = br.ReadBytes(2);
                    Offset = br.ReadInt32();
                }
                else
                {
                    throw new InvalidDataException($"Invalid BitMapType. Memorystream is not valid bitmap file");
                }
            }
        }

        internal BitMapFileHeader(BinaryReader br)
        {
                if (IsBmp(br))
                {
                    Size = br.ReadInt32();
                    Reserved1 = br.ReadBytes(2);
                    Reserved2 = br.ReadBytes(2);
                    Offset = br.ReadInt32();
                }
                else
                {
                    throw new InvalidDataException($"Invalid BitMapType. Memorystream is not valid bitmap file");
                }
        }

        protected bool IsBmp(BinaryReader br)
        {
            br.BaseStream.Seek(0, SeekOrigin.Begin);
            var signatureValue = br.ReadUInt16();

            if (Enum.IsDefined(typeof(BitMapType), signatureValue))
            {
                Signature = (BitMapType)signatureValue;
                return true;
            }

            return false;
        }

        internal void WriteBytes(BinaryWriter bw)
        {
            bw.Write((UInt16)Signature);
            bw.Write(Size);
            bw.Write(Reserved1);
            bw.Write(Reserved2);
            bw.Write(Offset);
        }
    }
}
