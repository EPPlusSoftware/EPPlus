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
    public class MarkToBaseTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void MarkToBaseTest()
        {
            var font = OpenTypeFonts.GetFontData(null, "Roboto", FontSubFamily.Regular, true);
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

                Assert.IsTrue(shaped.Glyphs.Any(x => x.YOffset > 0),
                    $"Expected YOffset > 0. Got: {string.Join(", ", shaped.Glyphs.Select(g => $"Y={g.YOffset}"))}");
            }
        }
    }
}
