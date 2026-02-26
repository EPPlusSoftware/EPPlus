using Microsoft.ApplicationInsights.DataContracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Reading
{
    [TestClass]
    public class VerticalTextTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }


        [TestMethod]
        public void TestVerticalMetrics_Vhea()
        {
            var font = OpenTypeFonts.LoadFont("BIZ UDGothic");
            Assert.IsNotNull(font);
            var vhea = font.VheaTable;
            Assert.IsNotNull(vhea, "vhea table should be present in a CJK font.");
            // Basic sanity checks on vhea values
            Assert.IsTrue(vhea.Ascent > 0, "Ascent should be positive.");
            Assert.IsTrue(vhea.AdvanceHeightMax > 0, "AdvanceHeightMax should be positive.");
            Assert.IsTrue(
               vhea.NumberOfVMetrics <= font.MaxpTable.numGlyphs,
               $"NumberOfVMetrics ({vhea.NumberOfVMetrics}) exceeds numGlyphs ({font.MaxpTable.numGlyphs}).");
        }

        [TestMethod]
        public void TestVerticalMetrics_Vmtx()
        {
            var font = OpenTypeFonts.LoadFont("BIZ UDGothic");
            Assert.IsNotNull(font);

            var vhea = font.VheaTable;
            var vmtx = font.VmtxTable;
            Assert.IsNotNull(vmtx, "vmtx table should be present in a CJK font.");

            // VMetrics count must match vhea.NumberOfVMetrics
            Assert.AreEqual(
                (int)vhea.NumberOfVMetrics,
                vmtx.VMetrics.Count,
                $"VMetrics.Count ({vmtx.VMetrics.Count}) should match vhea.NumberOfVMetrics ({vhea.NumberOfVMetrics}).");

            // Total glyph coverage must equal maxp.numGlyphs
            int totalCoverage = vmtx.VMetrics.Count + vmtx.TopSideBearings.Count;
            Assert.AreEqual(
                font.MaxpTable.numGlyphs,
                totalCoverage,
                $"VMetrics ({vmtx.VMetrics.Count}) + TopSideBearings ({vmtx.TopSideBearings.Count}) should equal numGlyphs ({font.MaxpTable.numGlyphs}).");

            // Majority of advance heights should be > 0 (some glyphs like .notdef may be 0)
            int zeroCount = 0;
            for (int i = 0; i < vmtx.VMetrics.Count; i++)
            {
                if (vmtx.VMetrics[i].AdvanceHeight == 0)
                    zeroCount++;
            }
            Assert.IsTrue(
                zeroCount < vmtx.VMetrics.Count / 2,
                $"Too many zero AdvanceHeights: {zeroCount} out of {vmtx.VMetrics.Count}.");

            // GetAdvanceHeight should work for first and last glyph
            Assert.IsTrue(vmtx.GetAdvanceHeight(0) > 0, "GetAdvanceHeight(0) should be > 0.");
            Assert.IsTrue(
                vmtx.GetAdvanceHeight((ushort)(font.MaxpTable.numGlyphs - 1)) > 0,
                "GetAdvanceHeight for last glyph should be > 0.");
        }
    }
}
