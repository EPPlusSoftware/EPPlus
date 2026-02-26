using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    [TestClass]
    public class MarkToBaseTests
    {
        [TestInitialize]
        public void TestSetup()
        {
            OpenTypeFonts.ClearFontCache();
            
        }

        [TestMethod]
        public void MarkToBaseTest()
        {
            var resolver = new DefaultFontResolver(null, true); // system-Roboto, ingen testmapp
            OpenTypeFonts.Configure(resolver);
            var font = OpenTypeFonts.LoadFont("Roboto", FontSubFamily.Regular, ignoreCache: true);
            Debug.WriteLine("=== MarkToBaseTest ===");
            Debug.WriteLine($"Font instance: {font.GetHashCode()}");
            Debug.WriteLine($"CmapTable instance: {font.CmapTable.GetHashCode()}");
            Debug.WriteLine($"SubTables count: {font.CmapTable.SubTables.Count}");
            for (int i = 0; i < font.CmapTable.SubTables.Count; i++)
                Debug.WriteLine($"  SubTable[{i}]: Format={font.CmapTable.SubTables[i].Format} HashCode={font.CmapTable.SubTables[i].GetHashCode()}");
            var shaper = new TextShaper(font);

            string test = "A\u0302\u0309";
            // Lägg till lite synchronization för att verifiera
            lock (typeof(MarkToBaseTests))
            {
                var shaped = shaper.Shape(test, ShapingOptions.Full);

                foreach (var g in shaped.Glyphs)
                {
                    Debug.WriteLine($"GID={g.GlyphId,-4} XAdv={g.XAdvance,-5} YOff={g.YOffset,-4}");
                }

                Debug.WriteLine($"GPOS null? {font.GposTable == null}");
                Debug.WriteLine($"FullyLoaded? {font.FullyLoaded}"); //

                Assert.IsTrue(shaped.Glyphs.Any(x => x.YOffset > 0),
                    $"Expected YOffset > 0. Got: {string.Join(", ", shaped.Glyphs.Select(g => $"Y={g.YOffset}"))}");
            }
        }
    }
}
