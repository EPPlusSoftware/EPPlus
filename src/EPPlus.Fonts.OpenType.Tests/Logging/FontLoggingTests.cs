using EPPlus.Fonts.OpenType.Logging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests
{
    /// <summary>
    /// The tests use no installed fonts: system directories are not searched, so the outcome is
    /// the same on every machine. Only the embedded fonts (Archivo Narrow, Noto Emoji) are used.
    /// Every test creates its own engine and logger, so they are safe to run in parallel.
    /// </summary>
    [TestClass]
    public class FontLoggingTests
    {
        private static OpenTypeFontEngine CreateEngine(IFontLogger logger)
        {
            return new OpenTypeFontEngine(c =>
            {
                c.SearchSystemDirectories = false;
                c.Logger = logger;
            });
        }

        private static FontLogEvent Single(CollectingFontLogger log, FontLogEventType type)
        {
            var events = log.GetEvents(type);
            Assert.AreEqual(1, events.Count, "Expected exactly one " + type + " event.");
            return events[0];
        }

        [TestMethod]
        public void UnknownFont_LogsLastResort()
        {
            var log = new CollectingFontLogger();
            using (var engine = CreateEngine(log))
            {
                Assert.IsNotNull(engine.GetTextShaper("NoSuchFont_EpplusLogTest"));
            }

            var e = Single(log, FontLogEventType.FontLastResort);
            Assert.AreEqual(FontLogSeverity.Warning, e.Severity);
            Assert.AreEqual("NoSuchFont_EpplusLogTest", e.RequestedFont);
            Assert.AreEqual("Archivo Narrow", e.ResolvedFont);
        }

        [TestMethod]
        public void LastResortFontRequestedExplicitly_LogsResolvedNotLastResort()
        {
            var log = new CollectingFontLogger();
            using (var engine = CreateEngine(log))
            {
                Assert.IsNotNull(engine.GetTextShaper("Archivo Narrow"));
            }

            Assert.AreEqual(0, log.GetEvents(FontLogEventType.FontLastResort).Count);
            Single(log, FontLogEventType.FontResolved);
        }

        [TestMethod]
        public void MissingGlyph_IsReportedOncePerCodePoint()
        {
            var log = new CollectingFontLogger();
            using (var engine = CreateEngine(log))
            {
                var font = engine.LoadFont("Archivo Narrow");
                var provider = new DefaultFontProvider(engine, font);
                OpenTypeFont used;
                ushort glyphId;

                // U+4F60 is Han. Archivo Narrow lacks it and no CJK font is available.
                Assert.IsFalse(provider.TryGetGlyphFont(0x4F60, out used, out glyphId));
                Assert.IsFalse(provider.TryGetGlyphFont(0x4F60, out used, out glyphId));
            }

            var missing = Single(log, FontLogEventType.GlyphMissing);
            Assert.AreEqual(FontLogSeverity.Warning, missing.Severity);
            Assert.AreEqual(0x4F60u, missing.CodePoint.Value);
            Assert.AreEqual(UnicodeScript.Han, missing.Script.Value);
            StringAssert.Contains(missing.Message, "U+4F60");

            // The chain is resolved once, and reported once, however many glyphs ask for it.
            var chain = Single(log, FontLogEventType.ScriptChainResolved);
            Assert.AreEqual(UnicodeScript.Han, chain.Script.Value);
            StringAssert.Contains(chain.Message, "Microsoft YaHei [not found]");
        }

        [TestMethod]
        public void EmojiFallback_IsReportedOnFirstUseOfTheFont()
        {
            var log = new CollectingFontLogger();
            using (var engine = CreateEngine(log))
            {
                var font = engine.LoadFont("Archivo Narrow");
                var provider = new DefaultFontProvider(engine, font);
                OpenTypeFont used;
                ushort glyphId;

                Assert.IsTrue(provider.TryGetGlyphFont(0x1F600, out used, out glyphId));
                Assert.IsTrue(provider.TryGetGlyphFont(0x1F601, out used, out glyphId));
            }

            var first = Single(log, FontLogEventType.ScriptFallbackUsed);
            Assert.AreEqual(0x1F600u, first.CodePoint.Value);
            StringAssert.Contains(first.ResolvedFont, "Noto");

            // Debug level also follows each code point routed to the fallback.
            Assert.AreEqual(2, log.GetEvents(FontLogEventType.GlyphFallback).Count);
        }

        [TestMethod]
        public void MinimumSeverity_FiltersDebugEvents()
        {
            var log = new CollectingFontLogger(FontLogSeverity.Warning);
            using (var engine = CreateEngine(log))
            {
                engine.GetTextShaper("NoSuchFont_EpplusLogTest");
                engine.GetTextShaper("Archivo Narrow");
            }

            Assert.IsTrue(log.Events.Count > 0);
            Assert.IsTrue(log.Events.All(e => e.Severity >= FontLogSeverity.Warning));
        }

        [TestMethod]
        public void LoggerThatThrows_DoesNotBreakFontResolution()
        {
            using (var engine = CreateEngine(new ThrowingLogger()))
            {
                var shaper = engine.GetTextShaper("NoSuchFont_EpplusLogTest");
                Assert.IsNotNull(shaper);
            }
        }

        [TestMethod]
        public void TextFileLogger_WritesOneLinePerEvent()
        {
            var path = Path.Combine(Path.GetTempPath(), "epplus_fontlog_" + Guid.NewGuid().ToString("N") + ".txt");
            try
            {
                var logger = FontLoggerFactory.CreateTextFileLogger(new FileInfo(path));
                using (var engine = CreateEngine(logger))
                {
                    engine.GetTextShaper("NoSuchFont_EpplusLogTest");
                }

                var lines = File.ReadAllLines(path);
                Assert.IsTrue(
                    lines.Any(l => l.StartsWith("WRN ") && l.Contains("NoSuchFont_EpplusLogTest")),
                    "Expected a warning line naming the unknown font.");
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        private sealed class ThrowingLogger : IFontLogger
        {
            public bool IsEnabled(FontLogSeverity severity)
            {
                return true;
            }

            public void Log(FontLogEvent logEvent)
            {
                throw new InvalidOperationException("The logger failed on purpose.");
            }
        }
    }
}