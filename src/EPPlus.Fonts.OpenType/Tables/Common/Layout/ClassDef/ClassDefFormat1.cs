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
namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.ClassDef
{
    public class ClassDefFormat1 : ClassDefTable
    {
        public ushort StartGlyphID { get; set; }
        public ushort GlyphCount { get; set; }
        public ushort[] ClassValueArray { get; set; }

        public ClassDefFormat1()
        {
            Format = 1;
        }

        public override int GetClass(ushort glyphId)
        {
            int index = glyphId - StartGlyphID;
            if (index < 0 || index >= GlyphCount)
                return 0;

            if (ClassValueArray == null || index >= ClassValueArray.Length)
                return 0;

            return ClassValueArray[index];
        }

        internal override void SerializeBody(FontsBinaryWriter writer)
        {
            writer.WriteUInt16BigEndian(StartGlyphID);
            writer.WriteUInt16BigEndian(GlyphCount);

            for (int i = 0; i < GlyphCount; i++)
            {
                ushort v = 0;
                if (ClassValueArray != null && i < ClassValueArray.Length)
                    v = ClassValueArray[i];

                writer.WriteUInt16BigEndian(v);
            }
        }
    }
}
