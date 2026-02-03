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
using EPPlus.Fonts.OpenType.Tables.Common.Layout.ClassDef;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2
{
    /// <summary>
    /// PairPos Format 2: Class-based kerning.
    /// Uses ClassDef1/ClassDef2 and a class matrix.
    /// </summary>
    public class PairPosSubTableFormat2 : PairPosSubTable
    {
        /// <summary>
        /// Class definition for first glyph (ClassDef1)
        /// </summary>
        public ClassDefTable ClassDef1 { get; set; }

        /// <summary>
        /// Class definition for second glyph (ClassDef2)
        /// </summary>
        public ClassDefTable ClassDef2 { get; set; }

        /// <summary>
        /// Number of classes in ClassDef1
        /// </summary>
        public ushort Class1Count { get; set; }

        /// <summary>
        /// Number of classes in ClassDef2
        /// </summary>
        public ushort Class2Count { get; set; }

        /// <summary>
        /// Matrix [Class1Count, Class2Count] of pair value records
        /// </summary>
        public PairValueRecord[,] ClassMatrix { get; set; }

       public override bool TryGetPairAdjustment(
            ushort firstGlyph,
            ushort secondGlyph,
            out ValueRecord value1,
            out ValueRecord value2)
        {
            value1 = null;
            value2 = null;

            System.Diagnostics.Debug.WriteLine($"[Format2] Trying pair: {firstGlyph} + {secondGlyph}");

            if (Coverage == null || ClassDef1 == null || ClassDef2 == null || ClassMatrix == null)
            {
                System.Diagnostics.Debug.WriteLine($"  ✗ Missing data: Cov={Coverage != null}, CD1={ClassDef1 != null}, CD2={ClassDef2 != null}, Matrix={ClassMatrix != null}");
                return false;
            }

            // First glyph must be in coverage
            int coverageIndex = Coverage.GetGlyphIndex(firstGlyph);
            System.Diagnostics.Debug.WriteLine($"  Coverage index for {firstGlyph}: {coverageIndex}");

            if (coverageIndex < 0)
            {
                System.Diagnostics.Debug.WriteLine($"  ✗ First glyph {firstGlyph} not in coverage");
                return false;
            }

            int class1 = ClassDef1.GetClass(firstGlyph);
            int class2 = ClassDef2.GetClass(secondGlyph);

            System.Diagnostics.Debug.WriteLine($"  Classes: class1={class1}, class2={class2}");
            System.Diagnostics.Debug.WriteLine($"  Matrix bounds: Class1Count={Class1Count}, Class2Count={Class2Count}");

            if (class1 < 0 || class2 < 0)
            {
                System.Diagnostics.Debug.WriteLine($"  ✗ Negative class!");
                return false;
            }

            if (class1 >= Class1Count || class2 >= Class2Count)
            {
                System.Diagnostics.Debug.WriteLine($"  ✗ Class out of bounds!");
                return false;
            }

            var record = ClassMatrix[class1, class2];

            if (record == null)
            {
                System.Diagnostics.Debug.WriteLine($"  ✗ Matrix[{class1},{class2}] is null");
                return false;
            }

            System.Diagnostics.Debug.WriteLine($"  Matrix[{class1},{class2}] exists:");
            System.Diagnostics.Debug.WriteLine($"    Value1: {(record.Value1 != null ? $"XAdv={record.Value1.XAdvance}" : "null")}");
            System.Diagnostics.Debug.WriteLine($"    Value2: {(record.Value2 != null ? $"XAdv={record.Value2.XAdvance}" : "null")}");

            if (record.Value1 == null && record.Value2 == null)
            {
                System.Diagnostics.Debug.WriteLine($"  ✗ Both values are null");
                return false;
            }

            value1 = record.Value1;
            value2 = record.Value2;

            System.Diagnostics.Debug.WriteLine($"  ✓ SUCCESS! Returning XAdvance={value1?.XAdvance ?? 0}");
            return true;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long subtableStart = writer.BaseStream.Position;

            // Header
            writer.WriteUInt16BigEndian(SubtableFormat); // 2

            long coverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Coverage offset placeholder

            long classDef1OffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // ClassDef1 offset placeholder

            long classDef2OffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // ClassDef2 offset placeholder

            writer.WriteUInt16BigEndian(ValueFormat1);
            writer.WriteUInt16BigEndian(ValueFormat2);

            writer.WriteUInt16BigEndian(Class1Count);
            writer.WriteUInt16BigEndian(Class2Count);

            // Class1Record/Class2Record matrix
            for (int i = 0; i < Class1Count; i++)
            {
                for (int j = 0; j < Class2Count; j++)
                {
                    var record = ClassMatrix?[i, j];

                    if (record != null) 
                    { 
                        record.Value1?.Write(writer, ValueFormat1); 
                        record.Value2?.Write(writer, ValueFormat2); 
                    }
                    else
                    { 
                        // Write empty value records 
                        var empty = new ValueRecord(); 
                        empty.Write(writer, ValueFormat1); 
                        empty.Write(writer, ValueFormat2); 
                    }   
                }
            }

            // Coverage
            if (Coverage != null)
            {
                ushort coverageOffset = (ushort)(writer.BaseStream.Position - subtableStart);
                long resumePos = writer.BaseStream.Position;

                writer.BaseStream.Seek(coverageOffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(coverageOffset);
                writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

                Coverage.Serialize(writer);
            }

            // ClassDef1
            if (ClassDef1 != null)
            {
                ushort classDef1Offset = (ushort)(writer.BaseStream.Position - subtableStart);
                long resumePos = writer.BaseStream.Position;

                writer.BaseStream.Seek(classDef1OffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(classDef1Offset);
                writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

                ClassDef1.Serialize(writer);
            }

            // ClassDef2
            if (ClassDef2 != null)
            {
                ushort classDef2Offset = (ushort)(writer.BaseStream.Position - subtableStart);
                long resumePos = writer.BaseStream.Position;

                writer.BaseStream.Seek(classDef2OffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(classDef2Offset);
                writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

                ClassDef2.Serialize(writer);
            }
        }
    }
}
