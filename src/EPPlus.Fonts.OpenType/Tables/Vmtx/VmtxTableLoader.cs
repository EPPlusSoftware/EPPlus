/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           vmtx table implementation (vertical text support)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Vmtx;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Vmtx
{
    /// <summary>
    /// Loads the 'vmtx' (Vertical Metrics) table.
    /// Requires vhea.numberOfVMetrics and maxp.numGlyphs to determine the table layout,
    /// analogous to how HmtxTableLoader uses hhea.numberOfHMetrics.
    /// </summary>
    internal class VmtxTableLoader : TableLoader<VmtxTable>
    {
        private readonly int _numberOfVMetrics;
        private readonly int _numGlyphs;

        public VmtxTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Vmtx)
        {
            _numberOfVMetrics = TableLoaders.GetVheaTableLoader(settings).Load().NumberOfVMetrics;
            _numGlyphs = TableLoaders.GetMaxpTableLoader(settings).Load().numGlyphs;
        }

        protected override VmtxTable LoadInternal()
        {
            _reader.BaseStream.Position = _offset;

            // Read longVerMetric array (numberOfVMetrics entries)
            var vMetrics = new List<LongVerMetric>(_numberOfVMetrics);
            for (int i = 0; i < _numberOfVMetrics; i++)
            {
                vMetrics.Add(new LongVerMetric
                {
                    AdvanceHeight = _reader.ReadUInt16BigEndian(),
                    TopSideBearing = _reader.ReadInt16BigEndian()
                });
            }

            // Read additional topSideBearing entries for remaining glyphs.
            // Count = numGlyphs - numberOfVMetrics
            int extraTsbCount = _numGlyphs - _numberOfVMetrics;
            var topSideBearings = new List<short>(extraTsbCount > 0 ? extraTsbCount : 0);
            for (int i = 0; i < extraTsbCount; i++)
            {
                topSideBearings.Add(_reader.ReadInt16BigEndian());
            }

            return new VmtxTable(vMetrics, topSideBearings);
        }
    }
}