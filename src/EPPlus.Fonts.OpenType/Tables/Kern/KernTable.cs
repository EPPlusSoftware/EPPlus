/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Tables.Kern
{
    public class KernTable : FontTableBase
    {
        public ushort version { get; set; }

        public ushort nTables { get; set; }

        public KernSubTable[] SubTables { get; set; }

        public ushort NumberOfFormat0Tables { get; set; }

        internal override void SerializeInternal(FontsBinaryWriter writer)
        {

            writer.WriteUInt16BigEndian(version);
            writer.WriteUInt16BigEndian((ushort)SubTables.Length);

            foreach (var subTable in SubTables)
            {
                subTable.Serialize(writer);
            }
        }
    }
}
