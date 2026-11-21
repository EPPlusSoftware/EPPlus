using EPPlus.Fonts.OpenType.Subsetting;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    /// <summary>
    /// Tests for the stacks per code point (<see cref="WebFontPlan.GetFamilyNames(string, FontSubFamily, int)"/>)
    /// and for the script chains that follow a glyph fallback written by name.
    ///
    /// Hiragana is routed to BIZ UDGothic, the only Japanese test font. Its script chain also names
    /// Yu Gothic and Meiryo, which are not installed here.
    /// </summary>
    [TestClass]
    public class WebFontPlanCodePointStackTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const int LatinA = 'a';
        private const int HiraganaA = 0x3042; // あ
        private const int HiraganaI = 0x3044; // い

        private static OpenTypeFontEngine CreateEngine(bool embedFallback)
        {
            return new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
                cfg.SetScriptFallback(UnicodeScript.Hiragana, "BIZ UDGothic", "Yu Gothic", "Meiryo");
                if (!embedFallback)
                {
                    cfg.OnFontEmbedding(info => info.FontName.StartsWith("BIZ UDGothic", StringComparison.OrdinalIgnoreCase)
                        ? FontEmbeddingDecision.Skip
                        : FontEmbeddingDecision.Default);
                }
            });
        }

        private static WebFontPlan Plan(OpenTypeFontEngine engine, string text)
        {
            var planner = new WebFontSubsetPlanner(engine);
            planner.AddText("Roboto", FontSubFamily.Regular, text);
            return planner.Build();
        }

        [TestMethod]
        public void FallbackByName_CodePointsGetAStackStartingWithTheFallback()
        {
            using (var engine = CreateEngine(embedFallback: false))
            {
                var plan = Plan(engine, "ab あい");

                var fontStack = plan.GetFamilyNames("Roboto", FontSubFamily.Regular);
                var kana = plan.GetFamilyNames("Roboto", FontSubFamily.Regular, HiraganaA);

                CollectionAssert.AreEqual(new[] { "BIZ UDGothic", "Yu Gothic", "Meiryo" }, kana.Take(3).ToArray());
                CollectionAssert.IsSubsetOf(fontStack.ToArray(), kana.ToArray(), "The font's stack follows the fallback's.");
                Assert.AreSame(kana, plan.GetFamilyNames("Roboto", FontSubFamily.Regular, HiraganaI),
                    "Code points routed to the same fallback share one stack instance.");
            }
        }

        [TestMethod]
        public void FallbackByName_OtherCodePointsGetTheFontStack()
        {
            using (var engine = CreateEngine(embedFallback: false))
            {
                var plan = Plan(engine, "ab あい");
                var fontStack = plan.GetFamilyNames("Roboto", FontSubFamily.Regular);

                Assert.AreSame(fontStack, plan.GetFamilyNames("Roboto", FontSubFamily.Regular, LatinA));
                Assert.AreSame(fontStack, plan.GetFamilyNames("Roboto", FontSubFamily.Regular, ' '));
                Assert.AreSame(fontStack, plan.GetFamilyNames("Roboto", FontSubFamily.Regular, 'Z'), "A code point that was never added.");
            }
        }

        [TestMethod]
        public void FallbackByName_FontStackListsTheScriptChainAfterTheFallback()
        {
            using (var engine = CreateEngine(embedFallback: false))
            {
                var plan = Plan(engine, "ab あい");
                var fontStack = plan.GetFamilyNames("Roboto", FontSubFamily.Regular).ToList();

                Assert.IsTrue(fontStack[0].EndsWith("+Roboto"), string.Join(", ", fontStack));
                var fallback = fontStack.IndexOf("BIZ UDGothic");
                Assert.IsTrue(fallback > 0, string.Join(", ", fontStack));
                Assert.AreEqual("Yu Gothic", fontStack[fallback + 1]);
                Assert.AreEqual("Meiryo", fontStack[fallback + 2]);
                Assert.AreEqual(1, plan.Faces.Count, "Only Roboto is embedded.");
            }
        }

        [TestMethod]
        public void EmbeddedFallback_HasNoStackOfItsOwn()
        {
            using (var engine = CreateEngine(embedFallback: true))
            {
                var plan = Plan(engine, "ab あい");
                var fontStack = plan.GetFamilyNames("Roboto", FontSubFamily.Regular);

                Assert.AreSame(fontStack, plan.GetFamilyNames("Roboto", FontSubFamily.Regular, HiraganaA),
                    "A subset only has the code points routed to it, so the font's stack draws them right.");
                Assert.AreEqual(2, plan.Faces.Count);
                Assert.IsFalse(fontStack.Contains("Yu Gothic"), "No script chain follows an embedded fallback.");
            }
        }

        [TestMethod]
        public void NoEmbedding_EmbedsNothingAndStartsWithTheRequestedFamily()
        {
            using (var engine = CreateEngine(embedFallback: true))
            {
                var planner = new WebFontSubsetPlanner(engine, false);
                planner.AddText("Roboto", FontSubFamily.Regular, "ab あい");
                planner.AddText("Calibri", FontSubFamily.Regular, "ab");
                var plan = planner.Build();

                Assert.AreEqual(0, plan.Faces.Count);
                CollectionAssert.AreEqual(new[] { "Roboto", "BIZ UDGothic", "Yu Gothic", "Meiryo" },
                    plan.GetFamilyNames("Roboto", FontSubFamily.Regular).ToArray());

                // Calibri is not installed: no Archivo Narrow safety net, Calibri first, then its chain.
                var calibri = plan.GetFamilyNames("Calibri", FontSubFamily.Regular).ToList();
                Assert.AreEqual("Calibri", calibri[0]);
                CollectionAssert.DoesNotContain(calibri, "Archivo Narrow");
            }
        }

        [TestMethod]
        public void NoEmbedding_FallbackCodePointsStillGetAStackOfTheirOwn()
        {
            using (var engine = CreateEngine(embedFallback: true))
            {
                var planner = new WebFontSubsetPlanner(engine, false);
                planner.AddText("Roboto", FontSubFamily.Regular, "ab あい");
                var plan = planner.Build();

                var fontStack = plan.GetFamilyNames("Roboto", FontSubFamily.Regular);
                var kana = plan.GetFamilyNames("Roboto", FontSubFamily.Regular, HiraganaA);
                Assert.AreNotSame(fontStack, kana);
                Assert.AreEqual("BIZ UDGothic", kana[0]);
                Assert.AreSame(fontStack, plan.GetFamilyNames("Roboto", FontSubFamily.Regular, LatinA));
            }
        }

        [TestMethod]
        public void UnknownFont_ReturnsNull()
        {
            using (var engine = CreateEngine(embedFallback: false))
            {
                var plan = Plan(engine, "ab");
                Assert.IsNull(plan.GetFamilyNames("Arial", FontSubFamily.Regular, LatinA));
            }
        }
    }
}
