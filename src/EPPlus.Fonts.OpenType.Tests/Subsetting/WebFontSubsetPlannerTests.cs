using EPPlus.Fonts.OpenType.Subsetting;
using OfficeOpenXml.Interfaces.Fonts;
using System.Text.RegularExpressions;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    /// <summary>
    /// Tests for <see cref="WebFontSubsetPlanner"/>: which fonts are embedded as subsets in web
    /// output, what each subset contains, and the font-family stack per measured font.
    ///
    /// The engines search only the test font folder, so Calibri and the other Office fonts are
    /// never installed here.
    /// </summary>
    [TestClass]
    public class WebFontSubsetPlannerTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const string ChartIncreasing = "\U0001F4C8"; // 📈

        private static OpenTypeFontEngine CreateEngine(Action<IEpplusFontConfiguration>? configure = null)
        {
            return new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
                configure?.Invoke(cfg);
            });
        }

        private static WebFontPlan Plan(OpenTypeFontEngine engine, params (string Family, FontSubFamily SubFamily, string Text)[] texts)
        {
            var planner = new WebFontSubsetPlanner(engine);
            foreach (var t in texts)
                planner.AddText(t.Family, t.SubFamily, t.Text);
            return planner.Build();
        }

        private static WebFontFace SingleFace(WebFontPlan plan, string sourceFamily)
        {
            var faces = plan.Faces.Where(f => f.SourceFamily == sourceFamily).ToList();
            Assert.AreEqual(1, faces.Count, "Expected one subset of " + sourceFamily + ".");
            return faces[0];
        }

        private static OpenTypeFont Parse(WebFontFace face)
        {
            return new OpenTypeFont(face.FontData);
        }

        private static bool HasGlyph(OpenTypeFont font, int codePoint)
        {
            return font.CmapTable.TryGetGlyphId((uint)codePoint, out _);
        }

        private static void AssertSubsetName(WebFontFace face)
        {
            StringAssert.Matches(face.FamilyName, new Regex("^[A-Z]{6}\\+" + Regex.Escape(face.SourceFamily) + "$"));
        }

        // -----------------------------------------------------------------------------------------
        // Embeddable font
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void OpenFont_IsSubsetAndFirstInStack()
        {
            using (var engine = CreateEngine())
            {
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, "Hello"));

                var face = SingleFace(plan, "Roboto");
                AssertSubsetName(face);
                CollectionAssert.AreEqual(new[] { face.FamilyName, "Roboto" }, plan.GetFamilyNames("Roboto", FontSubFamily.Regular).ToArray());
                Assert.AreEqual(GenericFontFamily.SansSerif, plan.GetGenericFamily("Roboto", FontSubFamily.Regular));
            }
        }

        [TestMethod]
        public void OpenFont_SubsetRoundTrips_WithTheRoutedCodePointsOnly()
        {
            using (var engine = CreateEngine())
            {
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, "Hello"));

                var parsed = Parse(SingleFace(plan, "Roboto"));
                foreach (var c in "Helo")
                    Assert.IsTrue(HasGlyph(parsed, c), "Missing '" + c + "'.");
                Assert.IsFalse(HasGlyph(parsed, 'z'));
            }
        }

        [TestMethod]
        public void FamilyLookup_IsCaseInsensitive()
        {
            using (var engine = CreateEngine())
            {
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, "Hello"));
                Assert.IsNotNull(plan.GetFamilyNames("ROBOTO", FontSubFamily.Regular));
                Assert.IsNull(plan.GetFamilyNames("Roboto", FontSubFamily.Bold));
            }
        }

        [TestMethod]
        public void BoldAndItalic_SetDescriptors()
        {
            using (var engine = CreateEngine())
            {
                var plan = Plan(engine,
                    ("Roboto", FontSubFamily.Bold, "Hi"),
                    ("Roboto", FontSubFamily.Italic, "Hi"),
                    ("Roboto", FontSubFamily.Regular, "Hi"));

                var bold = plan.Faces.Single(f => f.SubFamily == FontSubFamily.Bold);
                var italic = plan.Faces.Single(f => f.SubFamily == FontSubFamily.Italic);
                var regular = plan.Faces.Single(f => f.SubFamily == FontSubFamily.Regular);

                Assert.AreEqual(700, bold.FontWeight);
                Assert.IsFalse(bold.IsItalic);
                Assert.AreEqual(400, italic.FontWeight);
                Assert.IsTrue(italic.IsItalic);
                Assert.AreEqual(400, regular.FontWeight);
                Assert.IsFalse(regular.IsItalic);
            }
        }

        [TestMethod]
        public void EmptyText_GivesStackWithoutFace()
        {
            using (var engine = CreateEngine())
            {
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, ""));

                Assert.AreEqual(0, plan.Faces.Count);
                CollectionAssert.AreEqual(new[] { "Roboto" }, plan.GetFamilyNames("Roboto", FontSubFamily.Regular).ToArray());
            }
        }

        // -----------------------------------------------------------------------------------------
        // Shared subsets
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void SameFontInTwoRequests_IsSubsetOnce_WithUnionOfCodePoints()
        {
            // "Calibri" is not installed; the user chain resolves it to Roboto, so both requests
            // are measured with Roboto.
            using (var engine = CreateEngine(cfg => cfg.FontFallbacks["Calibri"] = new[] { "Roboto" }))
            {
                var plan = Plan(engine,
                    ("Roboto", FontSubFamily.Regular, "abc"),
                    ("Calibri", FontSubFamily.Regular, "xyz"));

                var face = SingleFace(plan, "Roboto");
                var parsed = Parse(face);
                foreach (var c in "abcxyz")
                    Assert.IsTrue(HasGlyph(parsed, c), "Missing '" + c + "'.");

                CollectionAssert.AreEqual(new[] { face.FamilyName, "Roboto" }, plan.GetFamilyNames("Roboto", FontSubFamily.Regular).ToArray());
                var calibri = plan.GetFamilyNames("Calibri", FontSubFamily.Regular).ToArray();
                Assert.AreEqual(face.FamilyName, calibri[0]);
                Assert.AreEqual("Calibri", calibri[1]);
                Assert.AreEqual("Roboto", calibri[2]);
            }
        }

        [TestMethod]
        public void Callback_IsCalledOncePerFont()
        {
            var calls = new Dictionary<string, int>();
            using (var engine = CreateEngine(cfg =>
            {
                cfg.FontFallbacks["Calibri"] = new[] { "Roboto" };
                cfg.OnFontEmbedding(info =>
                {
                    calls.TryGetValue(info.FontName, out var n);
                    calls[info.FontName] = n + 1;
                    return FontEmbeddingDecision.Default;
                });
            }))
            {
                Plan(engine,
                    ("Roboto", FontSubFamily.Regular, "abc"),
                    ("Calibri", FontSubFamily.Regular, "xyz"));

                Assert.IsTrue(calls.Count > 0);
                foreach (var kvp in calls)
                    Assert.AreEqual(1, kvp.Value, "Callback called more than once for " + kvp.Key);
            }
        }

        // -----------------------------------------------------------------------------------------
        // Glyph fallback
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Emoji_IsInNotoEmojiSubset_AfterPrimary()
        {
            using (var engine = CreateEngine())
            {
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, "Sales " + ChartIncreasing));

                var roboto = SingleFace(plan, "Roboto");
                var emoji = plan.Faces.Single(f => f.SourceFamily.Contains("Emoji"));
                AssertSubsetName(emoji);

                CollectionAssert.AreEqual(new[] { roboto.FamilyName, emoji.FamilyName, "Roboto" },
                    plan.GetFamilyNames("Roboto", FontSubFamily.Regular).ToArray());

                Assert.IsTrue(HasGlyph(Parse(emoji), 0x1F4C8));
                Assert.IsFalse(HasGlyph(Parse(emoji), 'S'));
                Assert.IsFalse(HasGlyph(Parse(roboto), 0x1F4C8));
            }
        }

        [TestMethod]
        public void CodePointNoFontHas_IsLeftOut()
        {
            using (var engine = CreateEngine())
            {
                // U+E000 is private use; no test font and no bundled fallback has it.
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, "ab\uE000"));

                Assert.AreEqual(1, plan.Faces.Count);
                var parsed = Parse(SingleFace(plan, "Roboto"));
                Assert.IsTrue(HasGlyph(parsed, 'a'));
                Assert.IsFalse(HasGlyph(parsed, 0xE000));
            }
        }

        // -----------------------------------------------------------------------------------------
        // Font that may not be embedded: name first, then the safety net
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void SkippedFont_NoChain_SafetyNetIsArchivoNarrow()
        {
            using (var engine = CreateEngine(cfg =>
                cfg.OnFontEmbedding(info => info.FontName.StartsWith("Roboto") ? FontEmbeddingDecision.Skip : FontEmbeddingDecision.Default)))
            {
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, "Hello"));

                Assert.IsFalse(plan.Faces.Any(f => f.SourceFamily == "Roboto"));
                var archivo = SingleFace(plan, "Archivo Narrow");
                Assert.IsTrue(HasGlyph(Parse(archivo), 'H'));

                CollectionAssert.AreEqual(new[] { "Roboto", archivo.FamilyName },
                    plan.GetFamilyNames("Roboto", FontSubFamily.Regular).ToArray());
            }
        }

        [TestMethod]
        public void SkippedFont_WithChain_SafetyNetIsFirstEmbeddableChainFont()
        {
            using (var engine = CreateEngine(cfg =>
            {
                cfg.FontFallbacks["Roboto"] = new[] { "Not Installed", "Open Sans" };
                cfg.OnFontEmbedding(info => info.FontName.StartsWith("Roboto") ? FontEmbeddingDecision.Skip : FontEmbeddingDecision.Default);
            }))
            {
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, "Hello"));

                var openSans = SingleFace(plan, "Open Sans");
                Assert.IsFalse(plan.Faces.Any(f => f.SourceFamily == "Archivo Narrow"));
                CollectionAssert.AreEqual(new[] { "Roboto", openSans.FamilyName, "Not Installed", "Open Sans" },
                    plan.GetFamilyNames("Roboto", FontSubFamily.Regular).ToArray());
            }
        }

        [TestMethod]
        public void EmbedWhole_IsNotEmbedded()
        {
            using (var engine = CreateEngine(cfg =>
                cfg.OnFontEmbedding(info => info.FontName.StartsWith("Roboto") ? FontEmbeddingDecision.EmbedWhole : FontEmbeddingDecision.Default)))
            {
                var plan = Plan(engine, ("Roboto", FontSubFamily.Regular, "Hello"));

                Assert.IsFalse(plan.Faces.Any(f => f.SourceFamily == "Roboto"));
                Assert.AreEqual("Roboto", plan.GetFamilyNames("Roboto", FontSubFamily.Regular)[0]);
            }
        }

        [TestMethod]
        public void CffFont_IsNotEmbedded()
        {
            using (var engine = CreateEngine())
            {
                RequireFont(engine, "Source Sans 3");
                var plan = Plan(engine, ("Source Sans 3", FontSubFamily.Regular, "Hello"));

                Assert.IsFalse(plan.Faces.Any(f => f.SourceFamily == "Source Sans 3"));
                var archivo = SingleFace(plan, "Archivo Narrow");
                CollectionAssert.AreEqual(new[] { "Source Sans 3", archivo.FamilyName },
                    plan.GetFamilyNames("Source Sans 3", FontSubFamily.Regular).ToArray());
            }
        }

        // -----------------------------------------------------------------------------------------
        // Font not installed
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void MissingFont_MeasuredFromMetrics_IsWrittenByName()
        {
            using (var engine = CreateEngine(cfg => cfg.MetricsFallback = MetricsFallbackMode.WhenFontMissing))
            {
                if (engine.GetMeasurementShaper("Calibri").HasGlyphIds)
                    Assert.Inconclusive("Test setup: Calibri is expected to be measured from serialized metrics.");

                var plan = Plan(engine, ("Calibri", FontSubFamily.Regular, "Hello"));
                var archivo = SingleFace(plan, "Archivo Narrow");

                CollectionAssert.AreEqual(
                    new[] { "Calibri", archivo.FamilyName, "Carlito", "Liberation Sans", "Arial", "Helvetica" },
                    plan.GetFamilyNames("Calibri", FontSubFamily.Regular).ToArray());
            }
        }

        [TestMethod]
        public void MissingFont_WithoutMetrics_EmbedsArchivoNarrowAndWarns()
        {
            var logger = new CapturingLogger();
            using (var engine = CreateEngine(cfg =>
            {
                cfg.MetricsFallback = MetricsFallbackMode.Disabled;
                cfg.Logger = logger;
            }))
            {
                var plan = Plan(engine, ("No Such Font", FontSubFamily.Regular, "Hello"));

                var archivo = SingleFace(plan, "Archivo Narrow");
                CollectionAssert.AreEqual(new[] { archivo.FamilyName, "No Such Font" },
                    plan.GetFamilyNames("No Such Font", FontSubFamily.Regular).ToArray());
                Assert.IsTrue(logger.Events.Any(e => e.Severity == FontLogSeverity.Warning
                    && e.Type == FontLogEventType.EmbeddingDecision
                    && e.Message.Contains("last-resort")));
            }
        }

        [TestMethod]
        public void ArchivoNarrow_IsNeverWrittenByName()
        {
            using (var engine = CreateEngine(cfg => cfg.MetricsFallback = MetricsFallbackMode.Disabled))
            {
                var plan = Plan(engine,
                    ("No Such Font", FontSubFamily.Regular, "Hello"),
                    ("Archivo Narrow", FontSubFamily.Regular, "Hello"));

                foreach (var family in new[] { "No Such Font", "Archivo Narrow" })
                {
                    foreach (var name in plan.GetFamilyNames(family, FontSubFamily.Regular))
                        Assert.AreNotEqual("Archivo Narrow", name, "Archivo Narrow written by name in the stack for " + family);
                }
            }
        }

        // -----------------------------------------------------------------------------------------
        // Usage
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Build_Twice_Throws()
        {
            using (var engine = CreateEngine())
            {
                var planner = new WebFontSubsetPlanner(engine);
                planner.AddText("Roboto", FontSubFamily.Regular, "a");
                planner.Build();
                Assert.ThrowsExactly<InvalidOperationException>(() => planner.Build());
                Assert.ThrowsExactly<InvalidOperationException>(() => planner.AddText("Roboto", FontSubFamily.Regular, "b"));
            }
        }

        private sealed class CapturingLogger : IFontLogger
        {
            public List<FontLogEvent> Events { get; } = new List<FontLogEvent>();
            public bool IsEnabled(FontLogSeverity severity) => true;
            public void Log(FontLogEvent logEvent) => Events.Add(logEvent);
        }
    }
}