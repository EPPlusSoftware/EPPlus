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
    }
}
