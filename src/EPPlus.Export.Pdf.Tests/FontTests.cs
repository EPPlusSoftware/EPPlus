/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************/
using EPPlus.Export.Pdf;
using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.DocumentObjects.Fonts;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace EPPlus.Export.Pdf.Tests
{
    /// <summary>
    /// Tests for how font objects are assigned object numbers and added to the
    /// PDF document body in the embedded-font path.
    ///
    /// In embedded mode the referable font resource (/F1 on the page) must be the
    /// Type0 font dictionary, which points through /DescendantFonts to the CIDFont
    /// that carries the /W width array. The original code instead built a simple
    /// Type1 PdfFont with a /Widths array (and a broken object number), which both
    /// misassigned object numbers (breaking rendering in Edge) and produced an
    /// invalid /Widths entry once the object was actually written.
    /// </summary>
    [TestClass]
    public class FontTests
    {
        // Roboto-Regular.ttf lives in the Fonts subfolder of the test project and is
        // copied next to the test assembly at build time.
        private const string TestFontName = "Roboto";

        // Build a font engine that only sees the test project's Fonts folder, so the
        // test is deterministic and never picks up a system font on a dev machine or CI.
        private static OpenTypeFontEngine CreateEngine()
        {
            var fontsFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
            return new OpenTypeFontEngine(cfg =>
            {
                cfg.FontDirectories.Add(fontsFolder);
                cfg.SearchSystemDirectories = false;
            });
        }

        private static PdfPageSettings CreateSettings(OpenTypeFontEngine engine, bool embeddFonts)
        {
            var settings = new PdfPageSettings(engine);
            settings.EmbeddFonts = embeddFonts;
            return settings;
        }

        private static PdfDictionaries CreateDictionariesWithSingleFont(PdfPageSettings settings, OpenTypeFontEngine engine)
        {
            var dictionaries = new PdfDictionaries();

            // In the new model Fonts is populated during shaping (ShapeText creates the resource,
            // GidsAndCharMap fills gids + charmap), NOT by AddFont. Reproduce that end state directly
            // so AddFontData has a realistic embedded resource to emit, without running a full export.
            var font = engine.LoadFont(TestFontName, FontSubFamily.Regular);
            var key = new FontKey(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum());

            var resource = new PdfFontResource(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum(), 1, settings);
            resource.fontData = font;

            // Populate a few glyphs as shaping would, so the embedded path (CIDSet, font stream subset)
            // has real glyph ids to work with.
            ushort gid;
            foreach (var ch in "Hi")
            {
                if (font.CmapTable.TryGetGlyphId(ch, out gid) && gid != 0)
                {
                    resource.Gids.Add(gid);
                    if (!resource.charactermappings.ContainsKey(gid))
                        resource.charactermappings[gid] = ch.ToString();
                }
            }

            dictionaries.Fonts[key] = resource;
            return dictionaries;
        }

        /// <summary>
        /// In embedded mode the page's font resource (/F1) must reference the Type0
        /// font dictionary. The referenced object number must be a real slot in the
        /// document body, and the object in that slot must be a PdfType0FontDict —
        /// not a simple Type1 PdfFont (which would carry an invalid /Widths array).
        /// </summary>
        [TestMethod]
        public void AddFontData_Embedded_FontResourcePointsAtType0Dict()
        {
            using (var engine = CreateEngine())
            {
                var settings = CreateSettings(engine, true);
                var dictionaries = CreateDictionariesWithSingleFont(settings, engine);
                var docSettings = PdfDocumentSettings.From(settings);

                var excelPdf = new ExcelPdf();
                excelPdf.SetDocumentSettingsForTest(PdfDocumentSettings.From(settings));
                excelPdf.SetDictionariesForTest(dictionaries);
                excelPdf.SetDocumentSettingsForTest(docSettings);
                excelPdf.AddFontData();

                var fontResource = dictionaries.GetFont(settings, TestFontName, FontSubFamily.Regular);

                Assert.AreNotEqual(-1, fontResource.fontObjectNumber,
                    "The referable font object number was never assigned.");

                // /F1 must point at the Type0 font dictionary in embedded mode.
                Assert.AreEqual(
                    fontResource.type0FontObjectNumber,
                    fontResource.fontObjectNumber,
                    "In embedded mode /F1 must reference the Type0 font dict, but " +
                    "fontObjectNumber does not match type0FontObjectNumber.");

                // That object number must map to a real slot in the body.
                Assert.IsTrue(
                    fontResource.fontObjectNumber >= 1 &&
                    fontResource.fontObjectNumber <= excelPdf._document.Count,
                    "The font object number does not point at an object in the body.");

                // The object in that slot must be the Type0 font dictionary.
                var objectAtFontSlot = excelPdf._document[fontResource.fontObjectNumber - 1];
                Assert.IsInstanceOfType(
                    objectAtFontSlot,
                    typeof(PdfType0FontDict),
                    "The slot /F1 references is occupied by " +
                    objectAtFontSlot.GetType().Name + ", not a PdfType0FontDict. A " +
                    "Type1 PdfFont here would emit an invalid /Widths array.");
            }
        }

        /// <summary>
        /// No simple Type1 PdfFont object may be emitted in embedded mode. Such an
        /// object would carry /Widths + /FirstChar + /LastChar, which is invalid for
        /// an embedded CID-keyed font and is what triggered the "invalid /Widths"
        /// error in PDF readers.
        /// </summary>
        [TestMethod]
        public void AddFontData_Embedded_DoesNotEmitSimpleFontObject()
        {
            using (var engine = CreateEngine())
            {
                var settings = CreateSettings(engine, true);
                var dictionaries = CreateDictionariesWithSingleFont(settings, engine);
                var docSettings = PdfDocumentSettings.From(settings);

                var excelPdf = new ExcelPdf();
                excelPdf.SetDocumentSettingsForTest(PdfDocumentSettings.From(settings));
                excelPdf.SetDictionariesForTest(dictionaries);
                excelPdf.SetDocumentSettingsForTest(docSettings);
                excelPdf.AddFontData();

                foreach (var obj in excelPdf._document)
                {
                    Assert.IsFalse(
                        obj is PdfFont,
                        "A simple Type1 PdfFont object was written in embedded mode; " +
                        "the referable font must be the Type0 font dict instead.");
                }
            }
        }

        [TestMethod]
        public void FontObjects_ForSubsetFont_ShareIdenticalTaggedName()
        {
            var engine = new OpenTypeFontEngine(cfg => cfg.SearchSystemDirectories = true);
            if(engine.GetFontAvailability("Calibri") != FontAvailability.Exact)
            {
                Assert.Inconclusive("Calibri font is not available on this system. This test requires Calibri to be installed.");
            }
            var font = engine.LoadFont("Calibri");
            var subset = new SingleFontSubsetter().Subset(font, new HashSet<int> { 65, 66, 67 }, true);

            var resource = new PdfFontResource("Calibri", FontSubFamily.Regular, 1, new PdfPageSettings(engine));
            resource.fontData = subset;

            // Object numbers are arbitrary; they only need to be set in dependency order.
            var descriptor = resource.GetFontDescriptorObject(10).RenderDictionary();
            var cid = resource.GetCIDFontObject(11).RenderDictionary();
            var type0 = resource.GetType0FontDictObject(12).RenderDictionary();

            var rx = new Regex(@"/(?:BaseFont|FontName)\s+/(\S+)");
            var names = new[] { descriptor, cid, type0 }
                .Select(s => rx.Match(s).Groups[1].Value)
                .ToList();

            foreach (var n in names)
                Assert.IsTrue(Regex.IsMatch(n, "^[A-Z]{6}\\+[^+\\s]+$"), n);
            Assert.AreEqual(1, names.Distinct().Count(), string.Join(" | ", names.ToArray()));
        }
    }
}