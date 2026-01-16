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

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.ClassDef
{
    public class ClassDefFormat2 : ClassDefTable
    {
        public List<ClassRangeRecord> ClassRangeRecords { get; set; }

        public ClassDefFormat2()
        {
            Format = 2;
        }

        public override int GetClass(ushort glyphId)
        {
            if (ClassRangeRecords == null)
                return 0;

            foreach (var r in ClassRangeRecords)
            {
                if (glyphId >= r.StartGlyphID && glyphId <= r.EndGlyphID)
                    return r.Class;
            }

            return 0;
        }

        internal override void SerializeBody(FontsBinaryWriter writer)
        {
            ushort count = (ushort)(ClassRangeRecords?.Count ?? 0);
            writer.WriteUInt16BigEndian(count);

            if (ClassRangeRecords != null)
            {
                foreach (var r in ClassRangeRecords)
                {
                    writer.WriteUInt16BigEndian(r.StartGlyphID);
                    writer.WriteUInt16BigEndian(r.EndGlyphID);
                    writer.WriteUInt16BigEndian(r.Class);
                }
            }
        }
    }
}
