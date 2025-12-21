using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    public static class ValidationTestHelper
    {
        private static string _fontFolder = string.Empty;
        private static List<string> _fontFolders = new List<string>();
        private static bool _initialized = false;
        private static object _syncRoot = new object();

        public static string FontFolder => _fontFolder;
        public static List<string> FontFolders => _fontFolders;

        public static void ClassInitialize(TestContext testContext)
        {
            if (!_initialized)
            {
                lock (_syncRoot)
                {
                    if (!_initialized)
                    {
                        _fontFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
                        _fontFolders.Clear();
                        _fontFolders.Add(_fontFolder);
                        OpenTypeFonts.ClearFontCache();
                        _initialized = true;
                    }
                }
            }
        }

        /// <summary>
        /// Helper to validate a table - optional utility
        /// </summary>
        public static void AssertTableValid<TTable, TValidator>(
            TTable table,
            OpenTypeFont font,
            string tableName)
            where TTable : FontTableBase
            where TValidator : ITableValidator<TTable>, new()
        {
            var validator = new TValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(table, context);

            Assert.IsTrue(result.IsValid,
                $"{tableName} validation failed for a known good font.");
        }
    }
}