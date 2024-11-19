/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using System.IO;
using System.Diagnostics;

namespace OfficeOpenXml.Drawing.EMF
{
    [DebuggerDisplay("Type: {Type}, Size: {Size}")]
    internal class EMR_RECORD
    {
        internal RECORD_TYPES Type; //4
        internal uint Size;         //4
        internal byte[] data;       //This byte array is used for records not yet implemented to preserve data.
        internal long position = 0;

        internal EMR_RECORD() { }

        internal EMR_RECORD(BinaryReader br, uint TypeValue, bool readData = false)
        {
            position = br.BaseStream.Position - 4;
            Type = (RECORD_TYPES)TypeValue;
            Size = br.ReadUInt32();
            if (readData && Size > 8)
            {
                data = br.ReadBytes((int)Size - 8);
            }
        }

        internal virtual void WriteBytes(BinaryWriter bw)
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
