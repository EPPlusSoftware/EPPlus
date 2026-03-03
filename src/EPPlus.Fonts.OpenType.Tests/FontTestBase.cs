/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/21/2025         EPPlus Software AB           Test base class
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Tests.Helpers;

namespace EPPlus.Fonts.OpenType.Tests
{
    /// <summary>
    /// Base class for all font tests
    /// </summary>
    public abstract class FontTestBase
    {
        /// <summary>
        /// Test context will be set by MsTest
        /// </summary>
        public abstract TestContext? TestContext { get; set; }

        /// <summary>
        /// Gets the font folder path (for reading test fonts)
        /// </summary>
        protected static string FontFolder => FontDirectoriesTestHelper.FontFolder;

        /// <summary>
        /// Gets the list of font folder paths (for reading test fonts)
        /// </summary>
        protected static List<string> FontFolders => FontDirectoriesTestHelper.FontFolders;

        /// <summary>
        /// Gets whether test output path is available (false in CI/CD)
        /// </summary>
        protected static bool IsTestOutputAvailable => FontDirectoriesTestHelper.IsTestOutputAvailable;

        /// <summary>
        /// Saves a font to the test output folder (c:\epplusTest\Fonts\).
        /// Skips silently if running in CI/CD environment.
        /// Follows EPPlus pattern: SaveWorkbook(name, pck)
        /// </summary>
        /// <param name="fileName">Filename (e.g. "subset_Roboto_abc.ttf")</param>
        /// <param name="font">Font to save</param>
        /// <returns>FileInfo for saved file, or null if output not available</returns>
        protected static FileInfo SaveFont(string fileName, OpenTypeFont font)
        {
            return FontDirectoriesTestHelper.SaveFontToOutput(font, fileName);
        }

        /// <summary>
        /// Saves a font to the test output folder using the current test name as filename.
        /// Automatically appends suffix if provided.
        /// Skips silently if running in CI/CD environment.
        /// </summary>
        /// <param name="font">Font to save</param>
        /// <param name="suffix">Optional suffix to append (e.g., "fi", "ff")</param>
        /// <returns>FileInfo for saved file, or null if output not available</returns>
        /// <example>
        /// SaveFontForCurrentTest(subset);           // → "Subset_Ff_ShouldHaveFfLigature.ttf"
        /// SaveFontForCurrentTest(subset, "fi");     // → "Subset_CommonLigatures_ShouldWork_fi.ttf"
        /// </example>
        protected FileInfo? SaveFontForCurrentTest(OpenTypeFont font, string suffix = "")
        {
            if (!IsTestOutputAvailable)
                return null;

            var testName = TestContext?.TestName ?? "UnknownTest";
            var safeSuffix = string.IsNullOrWhiteSpace(suffix) ? "" : $"_{suffix}";
            var fileName = $"{testName}{safeSuffix}.ttf";

            return FontDirectoriesTestHelper.SaveFontToOutput(font, fileName);
        }


        /// <summary>
        /// Gets a FileInfo for an output file in a subdirectory.
        /// Creates subdirectory if needed.
        /// Follows EPPlus pattern: GetOutputFile(subPath, fileName)
        /// </summary>
        /// <param name="subPath">Subdirectory (e.g. "Roboto")</param>
        /// <param name="fileName">Filename (e.g. "subset_abc.ttf")</param>
        /// <returns>FileInfo or null if output not available</returns>
        protected static FileInfo GetOutputFile(string subPath, string fileName)
        {
            return FontDirectoriesTestHelper.GetOutputFile(subPath, fileName);
        }

        /// <summary>
        /// Checks if a font file exists in the output folder
        /// </summary>
        /// <param name="fileName">Filename to check</param>
        /// <returns>True if file exists, false if not or output unavailable</returns>
        protected static bool ExistsOutputFont(string fileName)
        {
            return FontDirectoriesTestHelper.ExistsOutputFont(fileName);
        }

        /// <summary>
        /// Deletes a font file from the output folder if it exists
        /// </summary>
        /// <param name="fileName">Filename to delete</param>
        protected static void DeleteOutputFont(string fileName)
        {
            FontDirectoriesTestHelper.DeleteOutputFont(fileName);
        }

        [TestInitialize]
        public void ClearAllCaches()
        {
            OpenTypeFonts.ClearFontCache();
            ConfigureResolver();
        }

        protected virtual void ConfigureResolver()
        {
            OpenTypeFonts.Configure(x => x.SetFontResolver(new DefaultFontResolver(FontFolders, false)));
        }

        [ClassInitialize(InheritanceBehavior.BeforeEachDerivedClass)]
        public static void BaseClassInitialize(TestContext context)
        {
            FontDirectoriesTestHelper.ClassInitialize(context);
        }

        /// <summary>
        /// Temporarily configures the font resolver for the duration of one test.
        /// Restores the default test resolver automatically via [TestCleanup].
        /// </summary>
        protected static void UseSystemFonts()
        {
            OpenTypeFonts.Configure(x => x.SetFontResolver(new DefaultFontResolver(null, true)));
        }

        protected static void UseFontFolders(IEnumerable<string> directories, bool searchSystemDirectories = false)
        {
            OpenTypeFonts.Configure(x => x.SetFontResolver(new DefaultFontResolver(directories, searchSystemDirectories)));
        }
    }
}