using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.DrawingRenderer.Tests
{
    /// <summary>
    /// Base class for renderer tests that lay out or measure text.
    ///
    /// Mirrors FontTestBase in EPPlus.Fonts.OpenType.Tests: tests use one of the exposed engines,
    /// TestFolderEngine (test fonts only) or SystemFontsEngine (test fonts and system fonts), and
    /// the render context built on it. The test fonts are linked from
    /// EPPlus.Fonts.OpenType.Tests/Fonts and copied to the TestFonts folder in the output
    /// directory (see the csproj).
    ///
    /// The engines and contexts are shared by all tests. Do not dispose them.
    /// </summary>
    public abstract class FontTestBase : TestBase
    {
        /// <summary>
        /// Folder holding the test fonts in the output directory.
        /// </summary>
        protected static string FontFolder
        {
            get { return Path.Combine(AppContext.BaseDirectory, "TestFonts"); }
        }

        // -----------------------------------------------------------------------------------------
        // Engines
        // -----------------------------------------------------------------------------------------

        private static readonly Lazy<OpenTypeFontEngine> _testFolderEngine =
            new Lazy<OpenTypeFontEngine>(() => new OpenTypeFontEngine(cfg =>
            {
                cfg.FontDirectories.Add(GetExistingFontFolder());
                cfg.SearchSystemDirectories = false;
            }));

        private static readonly Lazy<OpenTypeFontEngine> _systemFontsEngine =
            new Lazy<OpenTypeFontEngine>(() => new OpenTypeFontEngine(cfg =>
            {
                cfg.FontDirectories.Add(GetExistingFontFolder());
                cfg.SearchSystemDirectories = true;
            }));

        private static readonly Lazy<RenderContext> _testFolderContext =
            new Lazy<RenderContext>(() => new RenderContext(() => TestFolderEngine));

        private static readonly Lazy<RenderContext> _systemFontsContext =
            new Lazy<RenderContext>(() => new RenderContext(() => SystemFontsEngine));

        /// <summary>
        /// Engine that searches only the test font folder. Use this for tests that can rely on
        /// the fonts bundled with the font test project (BIZUDGothic, CrimsonText, EBGaramond,
        /// Mulish, NotoEmoji, NotoSansMath, Oi, OpenSans, PinyonScript, Roboto, SourceSans3,
        /// UnicaOne, Archivo Narrow).
        /// </summary>
        protected static OpenTypeFontEngine TestFolderEngine
        {
            get { return _testFolderEngine.Value; }
        }

        /// <summary>
        /// Engine that searches both the test font folder and the system font directories. Use
        /// this only in tests that need fonts not bundled with the tests (e.g. Aptos Narrow,
        /// Calibri), together with <see cref="RequireFont"/> so the test is Inconclusive on
        /// machines that lack the font.
        /// </summary>
        protected static OpenTypeFontEngine SystemFontsEngine
        {
            get { return _systemFontsEngine.Value; }
        }

        /// <summary>
        /// Render context over <see cref="TestFolderEngine"/>.
        /// </summary>
        protected static RenderContext TestFolderContext
        {
            get { return _testFolderContext.Value; }
        }

        /// <summary>
        /// Render context over <see cref="SystemFontsEngine"/>.
        /// </summary>
        protected static RenderContext SystemFontsContext
        {
            get { return _systemFontsContext.Value; }
        }

        /// <summary>
        /// Marks the test Inconclusive unless the font is available in the engine with the
        /// requested subfamily.
        /// </summary>
        protected static void RequireFont(
            OpenTypeFontEngine engine,
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular)
        {
            var avail = engine.GetFontAvailability(fontName, subFamily);
            if (avail != FontAvailability.Exact)
            {
                Assert.Inconclusive(
                    "Test requires " + fontName + " " + subFamily +
                    " which is not available (availability: " + avail + ").");
            }
        }

        private static string GetExistingFontFolder()
        {
            var folder = FontFolder;
            if (!Directory.Exists(folder))
                throw new DirectoryNotFoundException("Test font folder not found: " + folder);
            return folder;
        }
    }
}
