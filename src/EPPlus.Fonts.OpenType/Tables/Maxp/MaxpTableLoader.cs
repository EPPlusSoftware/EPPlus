using System.Collections.Generic;

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
namespace EPPlus.Fonts.OpenType.Tables.Maxp
{
    internal class MaxpTableLoader : TableLoader<MaxpTable>
    {
        public MaxpTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Maxp)
        {
        }

        protected override MaxpTable LoadInternal()
        {
            var pos = _reader.BaseStream.Position;
            var versionRawValue = _reader.ReadInt32BigEndian();
            var pos2 = _reader.BaseStream.Position;
            var nGlyphs = _reader.ReadUInt16BigEndian();
            var maxp = new MaxpTable
            {
                version = new Version16Dot16(versionRawValue),
                numGlyphs = nGlyphs
            };
            if(maxp.version.Major == 1)
            {
                maxp.maxPoints = _reader.ReadUInt16BigEndian();
                maxp.maxContours = _reader.ReadUInt16BigEndian();
                maxp.maxCompositePoints = _reader.ReadUInt16BigEndian();
                maxp.maxCompositeContours = _reader.ReadUInt16BigEndian();
                maxp.maxZones = _reader.ReadUInt16BigEndian();
                maxp.maxTwilightPoints = _reader.ReadUInt16BigEndian();
                maxp.maxStorage = _reader.ReadUInt16BigEndian();
                maxp.maxFunctionDefs = _reader.ReadUInt16BigEndian();
                maxp.maxInstructionDefs = _reader.ReadUInt16BigEndian();
                maxp.maxStackElements = _reader.ReadUInt16BigEndian();
                maxp.maxSizeOfInstructions = _reader.ReadUInt16BigEndian();
                maxp.maxComponentElements = _reader.ReadUInt16BigEndian();
                maxp.maxComponentDepth = _reader.ReadUInt16BigEndian();
            }
            return maxp;
        }
    }
}
