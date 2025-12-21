using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Serialization
{
    public static class SerializationTestHelper
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
                    if (!_initialized) // ← Double-check
                    {
                        _fontFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
                        _fontFolders.Clear();
                        _fontFolders.Add(_fontFolder);
                        OpenTypeFonts.ClearFontCache();
                        _initialized = true; // ← Sist
                    }
                }
            }
        }
    }
}
