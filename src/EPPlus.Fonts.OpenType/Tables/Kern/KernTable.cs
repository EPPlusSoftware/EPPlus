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
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Kern
{
    public class KernTable : FontTableBase
    {
        public override string Name => TableNames.Kern;

        public override bool IsEssentialTable => false;

        public ushort version { get; set; }
        public ushort numberOfFormat0Tables { get; set; }

        public List<KernSubTable> SubTables { get; set; } = new List<KernSubTable>();

        internal override void Clear()
        {
            SubTables.Clear();
            version = 0;
            numberOfFormat0Tables = 0;
        }

        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            writer.WriteUInt16BigEndian(version);
            writer.WriteUInt16BigEndian((ushort)SubTables.Count);

            foreach (var subTable in SubTables)
            {
                subTable.Serialize(writer);
            }
        }
    }
}
