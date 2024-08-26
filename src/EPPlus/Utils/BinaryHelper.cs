using System;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Utils
{
    static internal class BinaryHelper
    {
        static internal string GetStringAndUnicodeString(BinaryReader br, uint size, Encoding enc)
        {
            string s = GetString(br, size, enc);
            int reserved = br.ReadUInt16();
            uint sizeUC = br.ReadUInt32();
            string sUC = GetString(br, sizeUC, Encoding.Unicode);
            return sUC.Length == 0 ? s : sUC;
        }

        static internal string GetString(BinaryReader br, uint size, Encoding enc)
        {
            if (size > 0)
            {
                byte[] byteTemp = new byte[size];
                byteTemp = br.ReadBytes((int)size);
                return enc.GetString(byteTemp);
            }
            else
            {
                return "";
            }
        }

        static internal byte[] GetByteArray(string str, Encoding enc)
        {
            if(str == null)
                return new byte[0];
            return enc.GetBytes(str);
        }

        static internal string GetString(byte[] bytes, Encoding enc)
        {
            if (bytes == null || bytes.Length <= 0)
                return "";
            return enc.GetString(bytes);
        }

        static internal byte[] ConcatenateByteArrays(params byte[][] arrays)
        {
            int dataLength = 0;
            foreach (var arr in arrays)
            {
                dataLength += arr.Length;
            }
            byte[] dataArray = new byte[dataLength];
            int offset = 0;
            foreach (var arr in arrays)
            {
                Buffer.BlockCopy(arr, 0, dataArray, offset, arr.Length);
                offset += arr.Length;
            }
            return dataArray;
        }
    }
}
